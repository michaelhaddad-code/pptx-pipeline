"""
Step 4 (Row 2): INJECT
Applies resolved token values from the config directly into slide XML files in _raw/.
Must be run after update_config.py.
Usage: python inject.py [config_path] [library_dir]

Architecture:
- Reads slide XML as raw text (preserving namespace prefixes, XML declaration, etc.)
- Uses ElementTree ONLY for read-only DOM queries (finding shapes, reading text)
- All modifications are string operations on the raw XML text
- NEVER calls tree.write() — writes the modified string back directly
"""

import os
import json
import re
import shutil
import sys
from typing import Optional
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
import logging

logger = logging.getLogger(__name__)


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Shape-level tag local names we walk up to from cNvPr
SHAPE_TAGS = {"sp", "pic", "grpSp", "graphicFrame"}


# ── Helpers: read-only ET queries on a string ────────────────────────────

def _parse_xml_readonly(xml_str: str) -> ET.Element:
    """Parse XML string into an ET root for read-only queries."""
    return ET.fromstring(xml_str)


def _find_shape_by_id(root: ET.Element, shape_id: str) -> Optional[ET.Element]:
    """Find a cNvPr element by its @id attribute (works across namespaces)."""
    for elem in root.iter():
        local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if local == "cNvPr" and elem.get("id") == shape_id:
            return elem
    return None


def _get_shape_element(cnvpr: ET.Element, root: ET.Element) -> Optional[ET.Element]:
    """Walk up from a cNvPr element to its enclosing shape element."""
    parent_map = {child: parent for parent in root.iter() for child in parent}
    node = cnvpr
    while node in parent_map:
        node = parent_map[node]
        local = node.tag.split("}")[-1] if "}" in node.tag else node.tag
        if local in SHAPE_TAGS:
            return node
    return None


def _get_shape_texts(shape_node: ET.Element) -> list:
    """Return list of all <a:t> text contents in the shape."""
    a_t = f"{{{NS['a']}}}t"
    return [t.text or "" for t in shape_node.findall(f".//{a_t}")]


def _get_shape_text_length(shape_node: ET.Element) -> int:
    """Total character count of all text runs in the shape."""
    return sum(len(t) for t in _get_shape_texts(shape_node))


# ── Helpers: find shape XML span in raw string ───────────────────────────

def _find_shape_span(xml_str: str, shape_id: str) -> Optional[tuple]:
    """
    Find the start and end offsets of the shape element containing the
    cNvPr with the given @id in the raw XML string.
    Returns (start, end) or None.
    """
    # Find the cNvPr tag with this id
    # Pattern matches cNvPr with id="<shape_id>" (id can appear anywhere in attributes)
    cnvpr_pattern = re.compile(
        r'<[^>]*?cNvPr\b[^>]*?\bid\s*=\s*"' + re.escape(shape_id) + r'"[^>]*?/?>'
    )
    all_matches = list(cnvpr_pattern.finditer(xml_str))
    if not all_matches:
        return None

    if len(all_matches) > 1:
        logger.warning("    [!] Shape id=%s: found %d cNvPr matches in XML; injecting into first", shape_id, len(all_matches))

    m = all_matches[0]
    cnvpr_pos = m.start()

    # Walk backwards to find the enclosing shape open tag
    # We look for the nearest opening <p:sp, <p:pic, <p:grpSp, <p:graphicFrame
    # (or without namespace prefix)
    shape_open_pattern = re.compile(
        r'<([a-zA-Z0-9_]+:)?(' + '|'.join(SHAPE_TAGS) + r')\b[^>]*?>'
    )

    # Find ALL shape open tags and pick the closest one that actually encloses our cNvPr
    candidates = list(shape_open_pattern.finditer(xml_str[:cnvpr_pos]))
    candidates = [m for m in candidates if m.start() <= cnvpr_pos]

    # Try candidates from closest to furthest (reverse order)
    for best_start in reversed(candidates):
        start = best_start.start()
        prefix = best_start.group(1) or ""
        tag_name = best_start.group(2)
        full_tag = prefix + tag_name

        # Find the matching closing tag, handling nesting
        depth = 0
        open_pat = re.compile(r'<' + re.escape(full_tag) + r'[\s>/]')
        close_pat = re.compile(r'</' + re.escape(full_tag) + r'\s*>')

        pos = start
        found = False
        while pos < len(xml_str):
            next_open = open_pat.search(xml_str, pos)
            next_close = close_pat.search(xml_str, pos)

            if next_close is None:
                break

            if next_open is not None and next_open.start() < next_close.start():
                depth += 1
                pos = next_open.end()
            else:
                depth -= 1
                if depth == 0:
                    end = next_close.end()
                    # Verify cnvpr is actually inside this shape
                    if cnvpr_pos >= start and cnvpr_pos < end:
                        return (start, end)
                    found = True
                    break
                pos = next_close.end()

        if found:
            continue  # This candidate didn't contain our cNvPr, try the next one

    return None


# ── Full text replacement (string-based) ──────────────────────────────────

def _escape_for_xml(text: str) -> str:
    """Escape text for safe insertion into XML content."""
    return xml_escape(text)


def _map_new_labels_to_runs(shape_xml: str, at_matches: list, new_value: str) -> list | None:
    """Map new text with label:value structure onto bold/body run pairs.

    Splits new_value by newlines, then for each line checks if it has a
    "label: value" pattern. If so, maps the label (with colon+space) onto
    the next available bold run and the value onto the next body run.

    Returns segments list or None if the text doesn't have label structure
    or there aren't enough bold/body run pairs.
    """
    rpr_pattern = re.compile(r'<a:rPr[^>]*/?>')
    n_runs = len(at_matches)

    # Classify each run as bold or not
    run_bold = []
    defrpr_pattern = re.compile(r'<a:defRPr[^>]*/?>')
    for am in at_matches:
        preceding = shape_xml[max(0, am.start() - 500):am.start()]
        rpr_m = list(rpr_pattern.finditer(preceding))
        is_bold = False
        if rpr_m:
            rpr_text = rpr_m[-1].group(0)
            # Check b="1" or b="true" (case-insensitive)
            is_bold = bool(re.search(r'\bb\s*=\s*"(1|true)"', rpr_text, re.IGNORECASE))
        if not is_bold:
            # Check for inherited bold from <a:defRPr> in the paragraph
            defrpr_m = list(defrpr_pattern.finditer(preceding))
            if defrpr_m:
                is_bold = bool(re.search(r'\bb\s*=\s*"(1|true)"', defrpr_m[-1].group(0), re.IGNORECASE))
        run_bold.append(is_bold)

    # Find bold/body pairs: a bold run followed by a non-bold run
    pairs = []
    i = 0
    while i < n_runs:
        if run_bold[i]:
            body_idx = i + 1 if (i + 1 < n_runs and not run_bold[i + 1]) else None
            pairs.append((i, body_idx))
            i = i + 2 if body_idx is not None else i + 1
        else:
            i += 1

    if not pairs:
        return None

    # Split new_value into lines, then parse label:value from each
    lines = new_value.split('\n')
    label_value_pairs = []
    for line in lines:
        colon_pos = line.find(':')
        if colon_pos != -1 and colon_pos < 60:
            label = line[:colon_pos + 1] + ' '
            value = line[colon_pos + 1:].strip()
            label_value_pairs.append((label, value))
        elif line.strip():
            # No colon — treat entire line as body text
            label_value_pairs.append(('', line.strip()))

    if not label_value_pairs:
        return None

    # Map label/value pairs onto bold/body run pairs
    segments = [''] * n_runs
    for j, (label, value) in enumerate(label_value_pairs):
        if j >= len(pairs):
            # More lines than run pairs — append to last body run
            last_bold_idx, last_body_idx = pairs[-1]
            target = last_body_idx if last_body_idx is not None else last_bold_idx
            segments[target] += '\n' + (label + value if label else value)
            continue
        bold_idx, body_idx = pairs[j]
        if label:
            segments[bold_idx] = label
            if body_idx is not None:
                segments[body_idx] = value
            else:
                segments[bold_idx] = label + value
        else:
            # No label, put in body run
            if body_idx is not None:
                segments[body_idx] = value
            else:
                segments[bold_idx] = value

    return segments


def _pick_body_run(shape_xml: str, at_matches: list) -> int:
    """Pick the best run index for fallback text placement.

    Prefers non-bold runs with the smallest font size (body text).
    Falls back to index 0 if no better candidate is found.
    """
    rpr_pattern = re.compile(r'<a:rPr[^>]*/?>')
    candidates = []
    for i, am in enumerate(at_matches):
        # Look backwards from the <a:t> tag to find the preceding <a:rPr>
        preceding = shape_xml[max(0, am.start() - 500):am.start()]
        rpr_m = list(rpr_pattern.finditer(preceding))
        if not rpr_m:
            candidates.append((i, False, 1100))  # no rPr = default formatting
            continue
        rpr = rpr_m[-1].group(0)
        bold = 'b="1"' in rpr
        sz_m = re.search(r'sz="(\d+)"', rpr)
        sz = int(sz_m.group(1)) if sz_m else 1100
        candidates.append((i, bold, sz))

    # Sort: non-bold first, then smallest font size
    candidates.sort(key=lambda c: (c[1], c[2]))
    return candidates[0][0] if candidates else 0


def _expand_newlines_to_paragraphs(shape_xml: str) -> str:
    """Split any <a:t> content containing literal newlines into separate <a:p> elements.

    For each <a:p> paragraph, if any <a:r> run's <a:t> contains '\\n', split that
    run at each newline, creating new <a:p> elements. Each new <a:p> gets the
    original paragraph's <a:pPr> and the run's <a:rPr>.
    """
    # Match full <a:p>...</a:p> paragraphs
    para_pattern = re.compile(
        r'(<[^>]*?:p\b[^>]*?>)(.*?)(</[^>]*?:p\s*>)',
        re.DOTALL
    )
    # Match <a:pPr .../> or <a:pPr ...>...</a:pPr>
    ppr_pattern = re.compile(
        r'<[^>]*?:pPr\b[^>]*?(?:/>|>.*?</[^>]*?:pPr\s*>)',
        re.DOTALL
    )
    # Match full <a:r>...</a:r> runs
    run_pattern = re.compile(
        r'(<[^>]*?:r\b[^>]*?>)(.*?)(</[^>]*?:r\s*>)',
        re.DOTALL
    )
    # Match <a:rPr .../> or <a:rPr ...>...</a:rPr>
    rpr_pattern = re.compile(
        r'<[^>]*?:rPr\b[^>]*?(?:/>|>.*?</[^>]*?:rPr\s*>)',
        re.DOTALL
    )
    # Match <a:t>...</a:t>
    at_pattern = re.compile(
        r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
        re.DOTALL
    )

    def _process_paragraph(para_m):
        p_open = para_m.group(1)
        p_inner = para_m.group(2)
        p_close = para_m.group(3)

        # Check if any run in this paragraph has a newline in its <a:t>
        has_newline = False
        for at_m in at_pattern.finditer(p_inner):
            if '\n' in at_m.group(2):
                has_newline = True
                break
        if not has_newline:
            return para_m.group(0)

        # Extract the paragraph properties
        ppr_m = ppr_pattern.search(p_inner)
        ppr_xml = ppr_m.group(0) if ppr_m else ""

        # Determine namespace prefix from the open tag (e.g., "a:" from "<a:p>")
        ns_m = re.match(r'<([^>]*?):p\b', p_open)
        ns_prefix = ns_m.group(1) + ":" if ns_m else ""

        # Collect all paragraph lines by splitting runs at newlines
        # Each "line" is a list of (rpr_xml, text) tuples
        lines = [[]]
        for run_m in run_pattern.finditer(p_inner):
            run_inner = run_m.group(2)
            rpr_rm = rpr_pattern.search(run_inner)
            rpr_xml = rpr_rm.group(0) if rpr_rm else ""
            at_rm = at_pattern.search(run_inner)
            if not at_rm:
                continue
            t_open = at_rm.group(1)
            t_text = at_rm.group(2)
            t_close = at_rm.group(3)
            if '\n' not in t_text:
                lines[-1].append((rpr_xml, t_open, t_text, t_close))
            else:
                parts = t_text.split('\n')
                for pi, part in enumerate(parts):
                    if pi > 0:
                        lines.append([])
                    lines[-1].append((rpr_xml, t_open, part, t_close))

        # Build new paragraphs
        paragraphs = []
        for line_runs in lines:
            runs_xml = ""
            for rpr_xml_item, t_open, text, t_close in line_runs:
                runs_xml += f"<{ns_prefix}r>"
                if rpr_xml_item:
                    runs_xml += rpr_xml_item
                runs_xml += t_open + text + t_close
                runs_xml += f"</{ns_prefix}r>"
            paragraphs.append(p_open + ppr_xml + runs_xml + p_close)

        return "".join(paragraphs)

    return para_pattern.sub(_process_paragraph, shape_xml)


def _replace_shape_text_paragraph_aware(shape_xml: str, new_value: str) -> tuple[str, bool] | tuple[None, bool]:
    """Paragraph-aware text replacement for multi-paragraph shapes.

    When the shape has multiple <a:p> paragraphs and the new text has \\n
    delimiters matching the paragraph count, process each paragraph
    independently. This preserves per-paragraph <a:pPr> (spacing, indent)
    and assigns label text to the correct bold+colored run even when the
    label text changes (e.g., 'Unified:' -> 'Managed:').

    Returns (modified_shape_xml, True) on success, or (None, False) if
    this strategy doesn't apply (caller should fall back).
    """
    new_paragraphs = new_value.split('\n')
    if len(new_paragraphs) < 2:
        return None, False

    # Find all <a:p>...</a:p> paragraphs in the shape
    para_pattern = re.compile(
        r'(<[^>]*?:p\b[^>]*?>)(.*?)(</[^>]*?:p\s*>)',
        re.DOTALL
    )
    at_pattern = re.compile(
        r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
        re.DOTALL
    )
    rpr_pattern = re.compile(r'<[^>]*?:rPr\b[^>]*?(?:/>|>.*?</[^>]*?:rPr\s*>)', re.DOTALL)

    para_matches = list(para_pattern.finditer(shape_xml))

    # Only use paragraph-aware mode when the shape has multiple paragraphs
    # with <a:r> runs containing text (skip empty/whitespace-only paragraphs)
    text_paras = []
    for pm in para_matches:
        runs_in_para = list(at_pattern.finditer(pm.group(0)))
        if runs_in_para:
            # Check if any run has non-whitespace text
            has_text = any(r.group(2).strip() for r in runs_in_para)
            if has_text:
                text_paras.append(pm)

    if len(text_paras) < 2 or len(new_paragraphs) != len(text_paras):
        return None, False

    # Process each paragraph independently
    new_shape_xml = shape_xml
    # Work backwards to preserve offsets
    for para_idx in range(len(text_paras) - 1, -1, -1):
        pm = text_paras[para_idx]
        para_xml = pm.group(0)
        new_para_text = new_paragraphs[para_idx]

        at_matches_in_para = list(at_pattern.finditer(para_xml))
        if not at_matches_in_para:
            continue

        n_runs = len(at_matches_in_para)

        # Classify runs as bold (label) or body within this paragraph
        run_is_bold = []
        for am in at_matches_in_para:
            preceding = para_xml[:am.start()]
            rpr_m = list(rpr_pattern.finditer(preceding))
            is_bold = False
            if rpr_m:
                rpr_text = rpr_m[-1].group(0)
                is_bold = bool(re.search(r'\bb\s*=\s*"(1|true)"', rpr_text, re.IGNORECASE))
            run_is_bold.append(is_bold)

        # Split the new paragraph text into label + body based on colon
        colon_pos = new_para_text.find(':')
        has_label = colon_pos != -1 and colon_pos < 60

        # Build segments for this paragraph's runs
        segments = [''] * n_runs

        if has_label and any(run_is_bold):
            # Put label (up to and including ': ') in the first bold run
            label_part = new_para_text[:colon_pos + 1] + ' '
            body_part = new_para_text[colon_pos + 1:].strip()

            bold_idx = run_is_bold.index(True)
            segments[bold_idx] = label_part

            # Put body text in the first non-bold run after the bold run
            body_idx = None
            for i in range(bold_idx + 1, n_runs):
                if not run_is_bold[i]:
                    body_idx = i
                    break
            if body_idx is not None:
                segments[body_idx] = body_part
            else:
                # No body run — append to bold run
                segments[bold_idx] = label_part + body_part
        else:
            # No label structure or no bold runs — put everything in the
            # first non-bold run (or first run if all bold)
            target_idx = 0
            for i in range(n_runs):
                if not run_is_bold[i]:
                    target_idx = i
                    break
            segments[target_idx] = new_para_text

        # Apply segments to this paragraph (work backwards within paragraph)
        new_para_xml = para_xml
        for j in range(n_runs - 1, -1, -1):
            am = at_matches_in_para[j]
            escaped_segment = _escape_for_xml(segments[j])
            new_at = am.group(1) + escaped_segment + am.group(3)
            new_para_xml = new_para_xml[:am.start()] + new_at + new_para_xml[am.end():]

        # Replace the paragraph in shape XML
        para_start = pm.start()
        para_end = pm.end()
        # Adjust for prior replacements — use absolute positions in new_shape_xml
        # Since we work backwards, positions are still valid
        new_shape_xml = new_shape_xml[:para_start] + new_para_xml + new_shape_xml[para_end:]

    logger.debug("    [path] paragraph-aware replacement, %d paragraphs", len(text_paras))
    return new_shape_xml, True


def _replace_shape_text(shape_xml: str, new_value: str) -> tuple[str, bool]:
    """
    Replace the text content of a shape with new_value, preserving run formatting.

    Strategy: use the original run texts as a structural template. For each
    original run, find the corresponding portion in the new text by matching
    the run's label/prefix pattern. This keeps each <a:rPr> (bold, italic,
    color, etc.) paired with the correct text segment.

    Falls back to simple single-run replacement only if new text has no
    structural overlap with the original runs.

    Returns (modified_shape_xml, was_modified).
    """
    # Try paragraph-aware replacement first for multi-paragraph shapes
    result, success = _replace_shape_text_paragraph_aware(shape_xml, new_value)
    if success:
        return result, True

    at_pattern = re.compile(
        r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
        re.DOTALL
    )
    at_matches = list(at_pattern.finditer(shape_xml))
    if not at_matches:
        return shape_xml, False

    n_runs = len(at_matches)

    # Extract original run texts (XML-unescaped for matching)
    original_texts = []
    for am in at_matches:
        raw = am.group(2)
        raw = raw.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        raw = raw.replace('&quot;', '"').replace('&apos;', "'")
        original_texts.append(raw)

    # Try structure-preserving replacement: match each original run's text
    # as a label/prefix in the new value to determine split points.
    # This works when the new text follows the same structure as the original
    # (e.g., "ACR Growth: <new body>\nConsumption Plan in MSX: <new body>").
    segments = _split_by_run_structure(original_texts, new_value)

    if segments is not None:
        logger.debug("    [path] structure-preserving replacement, segments=%s", segments)
    else:
        # Fallback: try to map new text's own label:value structure onto
        # the original bold/body run pairs. E.g. new text "name: michael\n
        # family name: haddad" maps "name: " to a bold run and "michael" to
        # a body run, reusing the original run formatting pattern.
        segments = _map_new_labels_to_runs(shape_xml, at_matches, new_value)
        if segments is not None:
            logger.debug("    [path] label-mapping replacement, segments=%s", segments)

    if segments is None:
        # Final fallback: keep ALL <a:r> elements intact, put full text in the
        # best body run's <a:t>, empty all other runs' <a:t>.
        # This preserves per-run <a:rPr> (bold, color, highlight) on every run.
        logger.debug("    [path] fallback: placing text in best body run, clearing others")
        best_idx = _pick_body_run(shape_xml, at_matches)
        segments = [''] * n_runs
        segments[best_idx] = new_value

    new_shape_xml = shape_xml
    for j in range(n_runs - 1, -1, -1):
        am = at_matches[j]
        escaped_segment = _escape_for_xml(segments[j])
        new_at = am.group(1) + escaped_segment + am.group(3)
        new_shape_xml = new_shape_xml[:am.start()] + new_at + new_shape_xml[am.end():]

    return new_shape_xml, True


def _split_by_run_structure(original_texts: list, new_value: str) -> list | None:
    """
    Split new_value into segments that map 1-to-1 onto original_texts runs.

    For each original run, looks for its label (the text up to and including
    the first colon or the full text if short) as a prefix marker in new_value.
    Text between markers is assigned to the preceding run.

    Returns list of segments (same length as original_texts), or None if
    the structure can't be matched.
    """
    n = len(original_texts)
    if n <= 1:
        return [new_value]

    # Identify "label" runs — short runs ending with colon+space, typically bold headers.
    # These are the structural anchors we match against in the new text.
    # Build pairs: (label_text, position_in_new_value)
    remaining = new_value
    segments = []
    matched = 0

    for i in range(n):
        orig = original_texts[i].strip()
        if not orig:
            segments.append('')
            continue

        # Try to find this run's text (or its label prefix) in the remaining text
        # For label runs (ending with ": " or ":"), match the label exactly
        # For body runs, take everything up to the next label
        label = orig
        # Check if this is a label run — has a colon with a short prefix before it.
        # e.g., "ACR Growth: ...", "CTA: ...", "Unified: ..." are all labels
        # regardless of how long the full run text is.
        colon_pos = orig.find(':')
        # Only treat as a label if the prefix before the colon also appears in new_value
        label_prefix_text = orig[:colon_pos] if colon_pos != -1 else ""
        is_label = (colon_pos != -1 and colon_pos < 60
                    and label_prefix_text and label_prefix_text in new_value)

        if is_label:
            # Find this label in the remaining new text
            pos = remaining.find(label)
            if pos == -1:
                # Try just the part before colon
                label_prefix = orig.split(':')[0] + ':'
                pos = remaining.find(label_prefix)
                if pos == -1:
                    # Can't find this label — structure doesn't match
                    if matched == 0:
                        return None
                    # Append rest to previous segment
                    segments[-1] += remaining
                    remaining = ''
                    segments.extend([''] * (n - len(segments)))
                    return segments
                # Found prefix — take up to end of prefix + trailing space
                end = pos + len(label_prefix)
                if end < len(remaining) and remaining[end] == ' ':
                    end += 1
                # Any text before this label belongs to previous run
                if pos > 0 and segments:
                    segments[-1] += remaining[:pos]
                segments.append(remaining[pos:end])
                remaining = remaining[end:]
                matched += 1
            else:
                # Found exact label
                end = pos + len(label)
                if end < len(remaining) and remaining[end] == ' ':
                    end += 1
                if pos > 0 and segments:
                    segments[-1] += remaining[:pos]
                segments.append(remaining[pos:end])
                remaining = remaining[end:]
                matched += 1
        else:
            # Body run — find where this run's content ends in the new text.

            # Short word-token runs (e.g., "above", "below") — match just the
            # word itself, preserving it as an isolated formatted token.
            if len(orig) <= 20 and ' ' not in orig.strip():
                word = orig.strip()
                wp = remaining.find(word)
                if wp != -1:
                    end = wp + len(word)
                    # Include trailing space if present (to match original spacing)
                    if end < len(remaining) and remaining[end] == ' ':
                        end += 1
                    segments.append(remaining[wp:end])
                    # Any text before the word belongs to previous segment
                    if wp > 0 and segments and len(segments) >= 2:
                        segments[-2] += remaining[:wp]
                    remaining = remaining[end:]
                    matched += 1
                    continue

            # Check if the next run's original text appears literally
            # in the remaining text (e.g., italic parenthetical). If so, use
            # that as the split point instead of the next label.
            next_split = len(remaining)

            if i + 1 < n:
                next_orig = original_texts[i + 1].strip()
                if next_orig:
                    fp = remaining.find(next_orig)
                    if fp != -1:
                        next_split = fp
                        segments.append(remaining[:next_split])
                        remaining = remaining[next_split:]
                        matched += 1
                        continue

            # Fall back: find the next label in the remaining text
            for future_i in range(i + 1, n):
                future_orig = original_texts[future_i].strip()
                future_colon = future_orig.find(':')
                if future_colon != -1 and future_colon < 60:
                    future_label = future_orig.split(':')[0] + ':'
                    fp = remaining.find(future_label)
                    if fp != -1:
                        # Check for preceding newline
                        nl = remaining.rfind('\n', 0, fp)
                        next_split = nl if nl != -1 else fp
                        break

            segment_text = remaining[:next_split]
            segments.append(segment_text)
            remaining = remaining[next_split:]
            matched += 1

    # Append any leftover text to the last segment
    if remaining.strip() and segments:
        segments[-1] += remaining
        remaining = ''

    # Pad or trim to match run count
    while len(segments) < n:
        segments.append('')
    if len(segments) > n:
        # Merge overflow into last segment
        segments = segments[:n - 1] + ['\n'.join(segments[n - 1:])]

    # Require at least 2 segments with actual text content.
    # If all text landed in one segment, the structure match failed — fall through.
    non_empty = sum(1 for s in segments if s.strip())
    return segments if non_empty >= 2 else None


# ── Legacy token replacement (string-based) ───────────────────────────────


def _replace_tokens_in_shape_xml(shape_xml: str, tokens: dict) -> tuple[str, int]:
    """
    Replace token strings within a shape's XML using string operations.
    Handles tokens split across consecutive <a:r> runs.
    Returns (modified_shape_xml, replacement_count).
    """
    replacements = 0

    for token, new_value in tokens.items():
        escaped_token = _escape_for_xml(token)
        escaped_value = _escape_for_xml(new_value)

        # Pass 1: single-run replacement — token is entirely within one <a:t>...</a:t>
        # Pattern to find <a:t ...>content</a:t> (with optional attributes like xml:space)
        at_pattern = re.compile(
            r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
            re.DOTALL
        )

        single_count = 0
        def _single_replace(m):
            nonlocal single_count
            open_tag, content, close_tag = m.group(1), m.group(2), m.group(3)
            if escaped_token in content:
                single_count += content.count(escaped_token)
                return open_tag + content.replace(escaped_token, escaped_value) + close_tag
            return m.group(0)

        shape_xml = at_pattern.sub(_single_replace, shape_xml)
        if single_count > 0:
            replacements += single_count
            continue

        # Pass 2: split-run replacement — token spans multiple <a:r> runs
        # Collect all <a:t> contents in order within each <a:p> paragraph
        para_pattern = re.compile(
            r'(<[^>]*?:p\b[^>]*?>)(.*?)(</[^>]*?:p\s*>)',
            re.DOTALL
        )

        for para_m in para_pattern.finditer(shape_xml):
            para_xml = para_m.group(0)

            # Extract all <a:t> elements with their positions within the paragraph
            at_matches = list(at_pattern.finditer(para_xml))
            if not at_matches:
                continue

            texts = [m.group(2) for m in at_matches]
            joined = "".join(texts)

            if escaped_token not in joined:
                continue

            # Found the token spanning runs — replace it
            replaced_joined = joined.replace(escaped_token, escaped_value)

            # Strategy: put all replacement text in the first contributing run,
            # empty the rest. We need to find which runs contribute to the token.
            # Simple approach: find the token in the joined text, determine which
            # runs are involved, collapse text into first, empty others.

            # Build a map of character positions to run indices
            char_to_run = []
            for i, t in enumerate(texts):
                char_to_run.extend([i] * len(t))

            # Find the token occurrence(s) in joined text
            search_pos = 0
            new_texts = list(texts)
            while True:
                idx = joined.find(escaped_token, search_pos)
                if idx == -1:
                    break

                token_end = idx + len(escaped_token)

                # Determine which runs are involved
                if idx < len(char_to_run):
                    first_run = char_to_run[idx]
                else:
                    break
                if token_end - 1 < len(char_to_run):
                    last_run = char_to_run[token_end - 1]
                else:
                    last_run = len(texts) - 1

                # Compute position within first run
                run_start = sum(len(texts[r]) for r in range(first_run))
                pos_in_first = idx - run_start

                # Build new text for the first run: prefix + replacement + suffix
                first_text = new_texts[first_run]
                # prefix is everything before the token starts in this run
                prefix = first_text[:pos_in_first]
                # How much of the token is in the first run
                first_run_end = sum(len(texts[r]) for r in range(first_run + 1))
                chars_in_first = min(token_end, first_run_end) - idx

                if first_run == last_run:
                    # Token is in single run in joined text (edge case)
                    suffix = first_text[pos_in_first + len(escaped_token):]
                    new_texts[first_run] = prefix + escaped_value + suffix
                else:
                    # Remove consumed part from first run, add replacement
                    suffix_first = ""  # the rest after token portion consumed
                    new_texts[first_run] = prefix + escaped_value

                    # Empty intermediate runs
                    for r in range(first_run + 1, last_run):
                        new_texts[r] = ""

                    # Handle last run: remove the consumed portion, keep the rest
                    last_run_start = sum(len(texts[r]) for r in range(last_run))
                    consumed_in_last = token_end - last_run_start
                    new_texts[last_run] = texts[last_run][consumed_in_last:]

                # Only handle first occurrence per pass, then re-read
                break

            # Rebuild the paragraph XML with new texts
            new_para_xml = para_xml
            for i in range(len(at_matches) - 1, -1, -1):
                m = at_matches[i]
                old_at = m.group(0)
                new_at = m.group(1) + new_texts[i] + m.group(3)
                # Replace this specific occurrence in the paragraph
                at_start = m.start()
                at_end = m.end()
                new_para_xml = new_para_xml[:at_start] + new_at + new_para_xml[at_end:]

            # Replace the paragraph in the shape XML
            shape_xml = shape_xml[:para_m.start()] + new_para_xml + shape_xml[para_m.end():]
            replacements += 1

            # After modifying, break and re-process (stale positions)
            # The outer loop over tokens will re-enter with fresh shape_xml
            break

    return shape_xml, replacements


# ── Font shrinking (string-based) ────────────────────────────────────────

_DEFAULT_SZ = 1800  # 18pt in hundredths of a point


def _apply_text_autofit(shape_xml: str, target_sz: int) -> str:
    """Scale run font sizes proportionally so the largest becomes target_sz.

    Finds the max sz across all runs in the shape, computes a scale factor
    (target_sz / max_sz), then scales each run's existing sz proportionally.
    Runs without an explicit sz are left untouched (they inherit from
    slide layout/master). This preserves size hierarchy (e.g. 28pt header
    + 10pt footnote).

    Args:
        shape_xml: raw XML string of the shape
        target_sz: desired font size for the largest run, in hundredths of a point

    Returns modified shape_xml.
    """
    rpr_pattern = re.compile(
        r'(<(?!/)[^>]*?:rPr\b)([^>]*?)(/?>)',
        re.DOTALL
    )
    sz_attr_pattern = re.compile(r'\bsz\s*=\s*"(\d+)"')
    ar_pattern = re.compile(
        r'(<[^>]*?:r\b[^>]*?>)(.*?)(</[^>]*?:r\s*>)',
        re.DOTALL
    )

    # First pass: find the max sz across all runs
    max_sz = 0
    for run_m in ar_pattern.finditer(shape_xml):
        for rpr_m in rpr_pattern.finditer(run_m.group(2)):
            sz_m = sz_attr_pattern.search(rpr_m.group(2))
            if sz_m:
                max_sz = max(max_sz, int(sz_m.group(1)))

    if max_sz == 0:
        return shape_xml  # no explicit sizes found, nothing to scale

    scale_factor = target_sz / max_sz

    def _scale_sz(m):
        prefix = m.group(1)
        attrs = m.group(2)
        close = m.group(3)
        sz_m = sz_attr_pattern.search(attrs)
        if sz_m:
            old_sz = int(sz_m.group(1))
            new_sz = max(int(round(old_sz * scale_factor)), 600)
            attrs = attrs[:sz_m.start()] + f'sz="{new_sz}"' + attrs[sz_m.end():]
        # If no sz attribute, leave untouched
        return prefix + attrs + close

    def _process_run(run_m):
        open_tag = run_m.group(1)
        content = run_m.group(2)
        close_tag = run_m.group(3)
        content = rpr_pattern.sub(_scale_sz, content)
        return open_tag + content + close_tag

    return ar_pattern.sub(_process_run, shape_xml)


def _shrink_font_in_shape_xml(shape_xml: str, original_len: int, new_len: int) -> str:
    """
    If injected text is longer than original, proportionally reduce font sizes
    in <a:rPr> elements within the shape XML using regex.
    Inserts a default sz when absent.
    """
    if new_len <= original_len or original_len == 0:
        return shape_xml

    scale = max(original_len / new_len, 0.75)

    # Pattern to find opening <a:rPr ...> tags only (not closing </a:rPr>)
    rpr_pattern = re.compile(
        r'(<(?!/)[^>]*?:rPr\b)([^>]*?)(/?>)',
        re.DOTALL
    )
    sz_attr_pattern = re.compile(r'\bsz\s*=\s*"(\d+)"')

    def _scale_rpr(m):
        prefix = m.group(1)
        attrs = m.group(2)
        close = m.group(3)

        sz_m = sz_attr_pattern.search(attrs)
        if sz_m:
            old_sz = int(sz_m.group(1))
            new_sz = max(int(round(old_sz * scale)), 600)
            attrs = attrs[:sz_m.start()] + f'sz="{new_sz}"' + attrs[sz_m.end():]
        # If no sz attribute, leave it alone — runs inherit from slide layout/master

        return prefix + attrs + close

    # Only modify rPr elements that are within <a:r> runs (not paragraph-level defRPr etc.)
    # We match a:r blocks and process rPr inside them
    # Only shrink runs containing injected text (non-empty <a:t>) — don't touch unmodified runs
    ar_pattern = re.compile(
        r'(<[^>]*?:r\b[^>]*?>)(.*?)(</[^>]*?:r\s*>)',
        re.DOTALL
    )
    at_content_pattern = re.compile(
        r'<[^>]*?:t\b[^>]*?>(.*?)</[^>]*?:t\s*>',
        re.DOTALL
    )

    def _process_run(run_m):
        open_tag = run_m.group(1)
        content = run_m.group(2)
        close_tag = run_m.group(3)
        # Only shrink if this run has non-empty <a:t> content (i.e., contains injected text)
        at_m = at_content_pattern.search(content)
        if at_m and at_m.group(1).strip():
            content = rpr_pattern.sub(_scale_rpr, content)
        return open_tag + content + close_tag

    return ar_pattern.sub(_process_run, shape_xml)


# ── Table cell data injection (string-based) ─────────────────────────────

def _inject_table_data(xml_str: str, shape: dict) -> tuple[str, int]:
    """
    Inject tabular data into a graphicFrame table shape.

    The shape's resolved_tokens should contain a single token whose value is a
    JSON string representing a list of row-dicts, e.g.:
        [{"Name": "Alice", "Score": "95"}, {"Name": "Bob", "Score": "88"}]

    Each dict key maps to a column header in the first table row.
    The function replaces cell text in existing data rows (skipping the header).
    All modifications are string-based; ET is used only for read-only queries.

    Returns (modified_xml_str, replacement_count).
    """
    shape_name = shape.get("shape_name", "")
    shape_id = str(shape.get("shape_id", ""))
    resolved_tokens = shape.get("resolved_tokens", {})

    # New path: resolved_value contains the JSON string directly
    raw_value = shape.get("resolved_value")
    if not raw_value:
        # Legacy fallback: extract from resolved_tokens
        resolved_tokens = shape.get("resolved_tokens", {})
        actual_tokens = _filter_resolved_tokens(resolved_tokens)
        if not actual_tokens:
            logger.info("    [skip] table %s (id=%s) -- no resolved values", shape_name, shape_id)
            return xml_str, 0
        raw_value = next(iter(actual_tokens.values()))

    try:
        rows_data = json.loads(raw_value)
    except (json.JSONDecodeError, TypeError) as exc:
        logger.warning("    [!] table %s (id=%s): failed to parse table JSON: %s", shape_name, shape_id, exc)
        return xml_str, 0

    if not isinstance(rows_data, list):
        logger.warning("    [!] table %s (id=%s): expected list of row dicts, got %s", shape_name, shape_id, type(rows_data).__name__)
        return xml_str, 0

    # Locate the shape span in raw XML
    span = _find_shape_span(xml_str, shape_id)
    if span is None:
        logger.warning("    [!] table %s (id=%s): could not find shape span in raw XML", shape_name, shape_id)
        return xml_str, 0

    s_start, s_end = span
    shape_xml = xml_str[s_start:s_end]

    # Verify an <a:tbl> exists inside the shape
    tbl_open = re.search(r'<[^>]*?:tbl\b[^>]*?>', shape_xml)
    if tbl_open is None:
        logger.warning("    [!] table %s (id=%s): no <a:tbl> element found in shape XML", shape_name, shape_id)
        return xml_str, 0

    # ── Adjust row count if data has more/fewer rows than template ──
    shape_xml = _adjust_table_rows(shape_xml, len(rows_data))
    # Re-splice adjusted shape back (row count may have changed)
    xml_str = xml_str[:s_start] + shape_xml + xml_str[s_end:]
    s_end = s_start + len(shape_xml)

    # ── Collect all <a:tr> row spans within shape_xml ──

    tr_open_pat = re.compile(r'<([a-zA-Z0-9_]+:)?tr\b[^>]*?>')
    at_pattern = re.compile(
        r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
        re.DOTALL
    )

    def _find_element_spans(xml, open_pat):
        """Find all top-level spans of elements matching open_pat in xml."""
        spans = []
        pos = 0
        while pos < len(xml):
            m_open = open_pat.search(xml, pos)
            if m_open is None:
                break
            pfx = m_open.group(1) or ""
            local_name = m_open.group(0).split(":")[0].lstrip("<") if not pfx else None
            full = pfx + "tr"
            depth = 1
            inner_pos = m_open.end()
            o_pat = re.compile(r'<' + re.escape(full) + r'[\s>/]')
            c_pat = re.compile(r'</' + re.escape(full) + r'\s*>')
            while inner_pos < len(xml) and depth > 0:
                n_open = o_pat.search(xml, inner_pos)
                n_close = c_pat.search(xml, inner_pos)
                if n_close is None:
                    break
                if n_open is not None and n_open.start() < n_close.start():
                    depth += 1
                    inner_pos = n_open.end()
                else:
                    depth -= 1
                    if depth == 0:
                        spans.append((m_open.start(), n_close.end()))
                    inner_pos = n_close.end()
            pos = inner_pos if inner_pos > m_open.end() else m_open.end()
        return spans

    tr_spans = _find_element_spans(shape_xml, tr_open_pat)

    if len(tr_spans) < 2:
        logger.warning("    [!] table %s (id=%s): need at least 2 rows (header + data), found %s", shape_name, shape_id, len(tr_spans))
        return xml_str, 0

    # ── Extract header column names via ET (read-only) ──
    # Wrap shape_xml with namespace declarations since it's a substring
    # of the full slide XML and doesn't carry the root's xmlns attrs.
    _ns_attrs = "".join(f' xmlns:{p}="{u}"' for p, u in NS.items())
    wrapped_xml = f"<_wrap{_ns_attrs}>{shape_xml}</_wrap>"
    root = _parse_xml_readonly(wrapped_xml)
    a_tr = f"{{{NS['a']}}}tr"
    a_tc = f"{{{NS['a']}}}tc"
    a_t = f"{{{NS['a']}}}t"

    tr_elements = root.findall(f".//{a_tr}")
    if not tr_elements:
        logger.warning("    [!] table %s (id=%s): ET found no <a:tr> elements", shape_name, shape_id)
        return xml_str, 0

    header_row = tr_elements[0]
    header_cells = header_row.findall(f".//{a_tc}")
    headers = []
    for cell in header_cells:
        cell_texts = [t.text or "" for t in cell.findall(f".//{a_t}")]
        headers.append("".join(cell_texts).strip())

    data_row_spans = tr_spans[1:]  # skip header row span

    if len(rows_data) > len(data_row_spans):
        logger.warning("    [warn] table %s (id=%s): data has %s rows but table only has %s data rows; extra data rows will be skipped", shape_name, shape_id, len(rows_data), len(data_row_spans))

    rows_to_fill = min(len(rows_data), len(data_row_spans))
    replacements = 0

    # ── Process rows in reverse order so earlier offsets stay valid ──

    tc_open_pat = re.compile(r'<([a-zA-Z0-9_]+:)?tc\b[^>]*?>')

    for i in range(rows_to_fill - 1, -1, -1):
        row_dict = rows_data[i]
        tr_start, tr_end = data_row_spans[i]
        row_xml = shape_xml[tr_start:tr_end]

        # Find <a:tc> spans within this row
        tc_spans = []
        tc_pos = 0
        while tc_pos < len(row_xml):
            mc = tc_open_pat.search(row_xml, tc_pos)
            if mc is None:
                break
            tc_prefix = mc.group(1) or ""
            tc_full = tc_prefix + "tc"
            tc_depth = 1
            tc_inner = mc.end()
            tc_o = re.compile(r'<' + re.escape(tc_full) + r'[\s>/]')
            tc_c = re.compile(r'</' + re.escape(tc_full) + r'\s*>')
            while tc_inner < len(row_xml) and tc_depth > 0:
                n_o = tc_o.search(row_xml, tc_inner)
                n_c = tc_c.search(row_xml, tc_inner)
                if n_c is None:
                    break
                if n_o is not None and n_o.start() < n_c.start():
                    tc_depth += 1
                    tc_inner = n_o.end()
                else:
                    tc_depth -= 1
                    if tc_depth == 0:
                        tc_spans.append((mc.start(), n_c.end()))
                    tc_inner = n_c.end()
            tc_pos = tc_inner if tc_inner > mc.end() else mc.end()

        # Replace cell contents in reverse column order to keep offsets valid
        for col_idx in range(min(len(tc_spans), len(headers)) - 1, -1, -1):
            header_name = headers[col_idx]
            if header_name not in row_dict:
                continue

            new_value = _escape_for_xml(str(row_dict[header_name]))
            tc_s, tc_e = tc_spans[col_idx]
            cell_xml = row_xml[tc_s:tc_e]

            # Replace all <a:t> text content in this cell
            # Strategy: put new value in first <a:t>, empty the rest
            at_matches = list(at_pattern.finditer(cell_xml))
            if not at_matches:
                continue

            new_cell_xml = cell_xml
            for j in range(len(at_matches) - 1, -1, -1):
                am = at_matches[j]
                if j == 0:
                    replacement_text = new_value
                else:
                    replacement_text = ""
                new_at = am.group(1) + replacement_text + am.group(3)
                new_cell_xml = new_cell_xml[:am.start()] + new_at + new_cell_xml[am.end():]

            row_xml = row_xml[:tc_s] + new_cell_xml + row_xml[tc_e:]
            replacements += 1

        shape_xml = shape_xml[:tr_start] + row_xml + shape_xml[tr_end:]

    # ── Apply geometry (row heights + font sizes) if layout was computed ──
    layout = shape.get("layout", {})
    computed = layout.get("_computed")
    if computed:
        shape_xml = _inject_table_geometry(shape_xml, computed)
        logger.info("    [ok] table %s (id=%s): geometry applied (row_h=%d, rows=%d%s)",
                     shape_name, shape_id, computed["row_height"], computed["total_rows"],
                     ", single-row" if computed.get("single_row") else "")

    # Splice modified shape_xml back into the full document
    xml_str = xml_str[:s_start] + shape_xml + xml_str[s_end:]

    if replacements > 0:
        logger.info("    [ok] table %s (id=%s): %s cell(s) injected across %s row(s)", shape_name, shape_id, replacements, rows_to_fill)
    else:
        logger.warning("    [!] table %s (id=%s): no cells matched header columns", shape_name, shape_id)

    return xml_str, replacements


# ── Table row adjustment (add/remove rows) ───────────────────────────────

def _find_element_spans_in(xml, open_pat):
    """Find all top-level spans of elements matching open_pat in xml."""
    spans = []
    pos = 0
    while pos < len(xml):
        m_open = open_pat.search(xml, pos)
        if m_open is None:
            break
        pfx = m_open.group(1) or ""
        full = pfx + "tr"
        depth = 1
        inner_pos = m_open.end()
        o_pat = re.compile(r'<' + re.escape(full) + r'[\s>/]')
        c_pat = re.compile(r'</' + re.escape(full) + r'\s*>')
        while inner_pos < len(xml) and depth > 0:
            n_open = o_pat.search(xml, inner_pos)
            n_close = c_pat.search(xml, inner_pos)
            if n_close is None:
                break
            if n_open is not None and n_open.start() < n_close.start():
                depth += 1
                inner_pos = n_open.end()
            else:
                depth -= 1
                if depth == 0:
                    spans.append((m_open.start(), n_close.end()))
                inner_pos = n_close.end()
        pos = inner_pos if inner_pos > m_open.end() else m_open.end()
    return spans


def _adjust_table_rows(shape_xml: str, n_data_rows: int) -> str:
    """Add or remove <a:tr> rows so the table has exactly 1 header + n_data_rows data rows.

    - If the table has fewer data rows than needed, clones the last data row.
    - If the table has more, removes excess rows from the end.
    - Cloned rows have their <a:t> text content cleared.
    Returns the modified shape_xml.
    """
    tr_open_pat = re.compile(r'<([a-zA-Z0-9_]+:)?tr\b[^>]*?>')
    tr_spans = _find_element_spans_in(shape_xml, tr_open_pat)

    if len(tr_spans) < 2:
        return shape_xml  # need at least header + 1 data row

    data_row_spans = tr_spans[1:]  # skip header
    n_current = len(data_row_spans)

    if n_data_rows == n_current:
        return shape_xml

    at_pattern = re.compile(
        r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
        re.DOTALL
    )

    if n_data_rows > n_current:
        # Clone the last data row
        last_start, last_end = data_row_spans[-1]
        template_row = shape_xml[last_start:last_end]

        # Clear text content in the cloned row
        def _clear_text(m):
            return m.group(1) + m.group(3)
        clean_row = at_pattern.sub(_clear_text, template_row)

        n_to_add = n_data_rows - n_current
        insert_point = last_end
        shape_xml = shape_xml[:insert_point] + (clean_row * n_to_add) + shape_xml[insert_point:]

    elif n_data_rows < n_current:
        # Remove excess rows from the end (process in reverse)
        for i in range(n_current - 1, n_data_rows - 1, -1):
            s, e = data_row_spans[i]
            shape_xml = shape_xml[:s] + shape_xml[e:]

    return shape_xml


# ── Table geometry injection (row heights + font sizes) ──────────────────

def _inject_table_geometry(shape_xml: str, computed: dict) -> str:
    """Apply computed row heights and font sizes to a table shape's XML.

    Args:
        shape_xml: raw XML string of the graphicFrame shape
        computed: dict from compute_table_layout with row_height, font_sizes, total_rows

    Returns modified shape_xml.
    """
    row_height = computed.get("row_height")
    font_sizes = computed.get("font_sizes", {})

    if not row_height:
        return shape_xml

    # ── Update <a:tr h="..."> attributes ──
    tr_h_pattern = re.compile(r'(<[^>]*?:tr\b[^>]*?\bh\s*=\s*")(\d+)(")')
    row_idx = [0]

    def _replace_row_h(m):
        row_idx[0] += 1
        return m.group(1) + str(row_height) + m.group(3)

    shape_xml = tr_h_pattern.sub(_replace_row_h, shape_xml)

    # ── Update <a:ext cy="..."> on the graphicFrame to match total table height ──
    total_rows = computed.get("total_rows", 0)
    if total_rows > 0:
        frame_padding = 216978  # from Lara's script
        new_cy = row_height * total_rows + frame_padding
        # Match the ext element that's a direct child of xfrm (the frame's own size)
        # This is the first <a:ext> in the graphicFrame, before the <a:graphic>
        ext_pattern = re.compile(
            r'(<[^>]*?:ext\b[^>]*?\bcx\s*=\s*"\d+"[^>]*?\bcy\s*=\s*")(\d+)(")'
        )
        # Only replace the FIRST occurrence (the frame's xfrm ext, not any inner ones)
        shape_xml = ext_pattern.sub(
            lambda m: m.group(1) + str(new_cy) + m.group(3),
            shape_xml, count=1
        )

    # ── Scale font sizes based on row type ──
    # Determine row boundaries to know which fonts to apply
    tr_open_pat = re.compile(r'<([a-zA-Z0-9_]+:)?tr\b[^>]*?>')
    tr_spans = _find_element_spans_in(shape_xml, tr_open_pat)

    rpr_pattern = re.compile(r'(<(?!/)[^>]*?:rPr\b)([^>]*?)(/?>)', re.DOTALL)
    sz_attr_pattern = re.compile(r'\bsz\s*=\s*"(\d+)"')

    for row_i, (rs, re_) in enumerate(tr_spans):
        if row_i == 0:
            target_sz = font_sizes.get("header", 1000)
        elif row_i == 1:
            target_sz = font_sizes.get("summary_row", 1700)
        else:
            target_sz = font_sizes.get("data_row", 1200)

        row_xml = shape_xml[rs:re_]

        def _scale_rpr_in_row(m):
            prefix = m.group(1)
            attrs = m.group(2)
            close = m.group(3)
            sz_m = sz_attr_pattern.search(attrs)
            if sz_m:
                attrs = attrs[:sz_m.start()] + f'sz="{target_sz}"' + attrs[sz_m.end():]
            else:
                attrs = attrs + f' sz="{target_sz}"'
            return prefix + attrs + close

        new_row_xml = rpr_pattern.sub(_scale_rpr_in_row, row_xml)
        shape_xml = shape_xml[:rs] + new_row_xml + shape_xml[re_:]

        # Adjust subsequent spans if length changed
        delta = len(new_row_xml) - len(row_xml)
        if delta != 0:
            tr_spans = [(s + delta if s > rs else s, e + delta if e > rs else e)
                        for s, e in tr_spans]

    return shape_xml


# ── Image geometry injection ──────────────────────────────────────────────

def _inject_image_geometry(xml_str: str, shape_id: str, computed: dict,
                           original_geometry: dict = None) -> str:
    """Modify a <p:pic> shape's xfrm to apply computed image dimensions.

    Logic:
    - Clear <a:srcRect> by replacing any existing one with zeroed attributes
    - Compute new_cy = original_cx * (img_height_px / img_width_px)
    - If new_cy <= original_cy: keep original_cy unchanged
    - If new_cy > original_cy: update ONLY the pic shape's <a:ext cy=>
    - If expand_height is False, skip cy expansion (contain-fit)
    - NEVER change x, y, or cx
    - NEVER touch any other shape on the slide
    - Also handle repositioning (auto-stack) via new_x/new_y

    Args:
        xml_str: full slide XML string
        shape_id: the cNvPr id of the pic shape
        computed: dict with cx, cy, offset_x, offset_y from compute_image_fit
        original_geometry: dict with original x, y from config (for offset calc)

    Returns modified xml_str.
    """
    span = _find_shape_span(xml_str, shape_id)
    if span is None:
        logger.warning("    [!] Image shape id=%s: could not find shape span", shape_id)
        return xml_str

    s_start, s_end = span
    shape_xml = xml_str[s_start:s_end]

    # Clear <a:srcRect> — replace any existing one with zeroed attributes
    srcrect_pattern = re.compile(r'<[^>]*?:srcRect\b[^>]*/>')
    srcrect_full_pattern = re.compile(r'<[^>]*?:srcRect\b[^>]*>.*?</[^>]*?:srcRect\s*>', re.DOTALL)
    if srcrect_full_pattern.search(shape_xml):
        shape_xml = srcrect_full_pattern.sub('<a:srcRect l="0" t="0" r="0" b="0"/>', shape_xml, count=1)
    elif srcrect_pattern.search(shape_xml):
        shape_xml = srcrect_pattern.sub('<a:srcRect l="0" t="0" r="0" b="0"/>', shape_xml, count=1)

    # Read original cx and cy from the shape XML
    ext_pattern = re.compile(
        r'(<[^>]*?:ext\b[^>]*?\bcx\s*=\s*")(\d+)("[^>]*?\bcy\s*=\s*")(\d+)(")'
    )
    ext_m = ext_pattern.search(shape_xml)
    if ext_m:
        original_cx = int(ext_m.group(2))
        original_cy = int(ext_m.group(4))
    else:
        original_cx = computed.get("cx", 0)
        original_cy = computed.get("cy", 0)

    # Compute new_cy from image aspect ratio if pixel dimensions are available
    img_width_px = computed.get("img_width_px", 0)
    img_height_px = computed.get("img_height_px", 0)

    final_cy = original_cy  # default: keep original
    if img_width_px and img_height_px:
        aspect_cy = int(round(original_cx * (img_height_px / img_width_px)))
        expand_height = computed.get("expand_height", True)
        if expand_height is not False and aspect_cy > original_cy:
            final_cy = aspect_cy
        elif aspect_cy < original_cy:
            # Shrink shape to match smaller image aspect ratio
            final_cy = aspect_cy
    else:
        # Fall back to computed cy if no pixel dimensions
        final_cy = computed.get("cy", original_cy)

    # Update ONLY cy, never cx — replace first <a:ext> occurrence
    if ext_m:
        shape_xml = ext_pattern.sub(
            lambda m: m.group(1) + m.group(2) + m.group(3) + str(final_cy) + m.group(5),
            shape_xml, count=1
        )

    # Update <a:off x="..." y="..."/> for repositioning (auto-stack) or offsets
    # NEVER change x or y except for explicit repositioning
    new_x = computed.get("new_x")
    new_y = computed.get("new_y")
    offset_x = computed.get("offset_x", 0)
    offset_y = computed.get("offset_y", 0)

    needs_position_update = (new_x is not None or new_y is not None or
                             ((offset_x or offset_y) and original_geometry))

    if needs_position_update:
        if new_x is not None:
            target_x = new_x
        elif original_geometry:
            target_x = original_geometry.get("x", 0) + offset_x
        else:
            target_x = None

        if new_y is not None:
            target_y = new_y
        elif original_geometry:
            target_y = original_geometry.get("y", 0) + offset_y
        else:
            target_y = None

        if target_x is not None and target_y is not None:
            off_pattern = re.compile(
                r'(<[^>]*?:off\b[^>]*?\bx\s*=\s*")(\d+)("[^>]*?\by\s*=\s*")(\d+)(")'
            )
            shape_xml = off_pattern.sub(
                lambda m: m.group(1) + str(target_x) + m.group(3) + str(target_y) + m.group(5),
                shape_xml, count=1
            )

    xml_str = xml_str[:s_start] + shape_xml + xml_str[s_end:]
    return xml_str


# ── Slide dimension constants ─────────────────────────────────────────────

SLIDE_WIDTH_EMU = 12192000   # Standard 10" (16:9) slide width
SLIDE_HEIGHT_EMU = 6858000   # Standard 7.5" slide height


# ── Consolidated image layout ─────────────────────────────────────────────

def _layout_image_sections(xml_str: str, images_with_geometry: list,
                           all_shapes: list) -> str:
    """Single-pass layout for stacked image sections after image file replacement.

    Replaces _auto_shift_below_image, _normalize_image_gaps,
    _sync_overlay_positions, _detect_and_resize_overlays, and _check_slide_bounds
    with one consolidated pass.

    STEP 1: Determine final cy for every image (aspect ratio at column width).
    STEP 2: One top-to-bottom layout pass — place each label then image using
            the final cy values, with uniform 50K gaps.
    STEP 3: Apply all position + size changes to XML in one bottom-to-top pass.
    """
    if not images_with_geometry:
        return xml_str

    SECTION_GAP = 50000   # gap before each label
    LABEL_IMG_GAP = 5000  # gap between label bottom and image top
    SLIDE_BOTTOM_MARGIN = 50000

    off_re = re.compile(
        r'<[^>]*?:off\b[^>]*?\bx\s*=\s*"(\d+)"[^>]*?\by\s*=\s*"(\d+)"'
    )
    ext_re = re.compile(
        r'<[^>]*?:ext\b[^>]*?\bcx\s*=\s*"(\d+)"[^>]*?\bcy\s*=\s*"(\d+)"'
    )
    half_slide = SLIDE_WIDTH_EMU // 2

    # ── Helper: read current geometry from XML ──
    def _read_xml_geo(sid):
        span = _find_shape_span(xml_str, sid)
        if not span:
            return None
        snippet = xml_str[span[0]:span[1]]
        off_m = off_re.search(snippet)
        ext_m = ext_re.search(snippet)
        if off_m and ext_m:
            return {
                "x": int(off_m.group(1)), "y": int(off_m.group(2)),
                "cx": int(ext_m.group(1)), "cy": int(ext_m.group(2)),
            }
        return None

    # ── STEP 1: Compute final cy for every dynamic image ──
    # Compare the replacement image's aspect ratio against the original
    # placeholder's aspect ratio.  If they're similar (within 5%), the image
    # is effectively the same shape — keep the original cy to avoid rounding-
    # induced relayout.  Only resize when the aspect ratio actually changed.
    ASPECT_TOLERANCE = 0.05  # 5% tolerance

    dynamic_ids = set()
    final_cys = {}  # shape_id -> final cy
    for img in images_with_geometry:
        sid = str(img["target_shape_id"])
        dynamic_ids.add(sid)
        computed = img["_computed"]
        geo = _read_xml_geo(sid)
        if not geo:
            continue
        img_w = computed.get("img_width_px", 0)
        img_h = computed.get("img_height_px", 0)
        if img_w and img_h and geo["cx"] and geo["cy"]:
            new_aspect = img_h / img_w
            orig_aspect = geo["cy"] / geo["cx"]
            if abs(new_aspect - orig_aspect) / max(orig_aspect, 1e-9) <= ASPECT_TOLERANCE:
                # Aspect ratios are close enough — keep original box size
                final_cys[sid] = geo["cy"]
            else:
                # Aspect ratio changed — resize to fit width proportionally
                final_cys[sid] = int(round(geo["cx"] * new_aspect))
        else:
            final_cys[sid] = computed.get("cy", geo["cy"])

    # ── Collect ALL images in the same column (including static) ──
    all_image_entries = []  # (sid, geo_from_xml)
    for s in all_shapes:
        sid = str(s.get("shape_id", ""))
        if s.get("category") == "image" or sid in dynamic_ids:
            geo = _read_xml_geo(sid)
            if geo:
                all_image_entries.append((sid, geo))

    if not all_image_entries:
        return xml_str

    # Determine column from first dynamic image
    ref_sid = str(images_with_geometry[0]["target_shape_id"])
    ref_geo = next((g for sid, g in all_image_entries if sid == ref_sid), None)
    if not ref_geo:
        return xml_str
    col_is_left = (ref_geo["x"] + ref_geo["cx"] // 2) < half_slide

    # Filter to same column, sort by y
    col_images = sorted(
        [(sid, geo) for sid, geo in all_image_entries
         if ((geo["x"] + geo["cx"] // 2) < half_slide) == col_is_left],
        key=lambda x: x[1]["y"]
    )

    if not col_images:
        return xml_str

    # For static images, final cy = current cy (unchanged)
    for sid, geo in col_images:
        if sid not in final_cys:
            final_cys[sid] = geo["cy"]

    # ── Guard: skip layout pass if no image actually changed size ──
    # Only run stacking/shifting when an image expanded or shrunk significantly.
    # A tolerance of 20000 EMU (~1.6pt) absorbs rounding differences between
    # pixel-based aspect-ratio math and the original EMU dimensions.
    CY_TOLERANCE = 20000
    any_size_changed = False
    for sid, geo in col_images:
        if sid in dynamic_ids:
            if abs(final_cys[sid] - geo["cy"]) > CY_TOLERANCE:
                any_size_changed = True
                break
    if not any_size_changed:
        logger.info("    [layout] no image size changes detected, skipping relayout")
        return xml_str

    # ── Build set of group-child shape ids to skip ──
    group_child_ids = {str(s.get("shape_id", "")) for s in all_shapes if s.get("parent_group")}

    # ── Identify overlay shapes (Content Placeholders sitting on images) ──
    image_y_set = {geo["y"] for _, geo in col_images}
    overlay_parent = {}  # overlay_sid -> image_sid
    for s in all_shapes:
        sid = str(s.get("shape_id", ""))
        if sid in dynamic_ids or sid in group_child_ids:
            continue
        geo = _read_xml_geo(sid)
        if not geo:
            continue
        # Overlay = shape whose y matches an image y (within tolerance)
        # and whose x range falls within the image's x range
        for img_sid, img_geo in col_images:
            if (abs(geo["y"] - img_geo["y"]) < 20000 and
                    geo["x"] >= img_geo["x"] - 10000 and
                    geo["x"] + geo["cx"] <= img_geo["x"] + img_geo["cx"] + 10000 and
                    sid != img_sid):
                overlay_parent[sid] = img_sid
                break

    # ── Find text labels between consecutive images ──
    labels_per_section = {}  # section_index -> [(sid, geo)]
    for s in all_shapes:
        sid = str(s.get("shape_id", ""))
        if s.get("category") != "text" or sid in overlay_parent or sid in group_child_ids:
            continue
        geo = _read_xml_geo(sid)
        if not geo:
            continue
        if any(abs(geo["y"] - iy) < 20000 for iy in image_y_set):
            continue
        s_center = geo["x"] + geo["cx"] // 2
        if (s_center < half_slide) != col_is_left:
            continue
        # Assign to section by original y order (between image tops)
        for i in range(len(col_images) - 1):
            _, g_top = col_images[i]
            _, g_bot = col_images[i + 1]
            if g_top["y"] - 10000 <= geo["y"] < g_bot["y"] + 10000:
                labels_per_section.setdefault(i, []).append((sid, geo))
                break

    # ── Find the anchor label (closest label above first image, same column) ──
    first_img_y = col_images[0][1]["y"]
    anchor_label = None
    for s in all_shapes:
        sid = str(s.get("shape_id", ""))
        if s.get("category") != "text" or sid in group_child_ids:
            continue
        geo = _read_xml_geo(sid)
        if not geo:
            continue
        s_center = geo["x"] + geo["cx"] // 2
        if (s_center < half_slide) != col_is_left:
            continue
        # Must be above the first image, and closest to it (highest y that's still above)
        if geo["y"] < first_img_y and geo["y"] + geo["cy"] <= first_img_y + 20000:
            if anchor_label is None or geo["y"] > anchor_label[1]["y"]:
                anchor_label = (sid, geo)

    # ── STEP 2: Compute available space and scale if needed ──
    anchor_y = anchor_label[1]["y"] if anchor_label else first_img_y
    available = SLIDE_HEIGHT_EMU - anchor_y - SLIDE_BOTTOM_MARGIN

    # Tally fixed heights (labels + gaps)
    total_label_h = 0
    total_gaps = 0
    # Anchor label
    if anchor_label:
        total_label_h += anchor_label[1]["cy"]
        total_gaps += LABEL_IMG_GAP
    # Inter-image labels
    for i in range(len(col_images) - 1):
        section_labels = labels_per_section.get(i, [])
        total_gaps += SECTION_GAP
        for _, lg in section_labels:
            total_label_h += lg["cy"]
            total_gaps += LABEL_IMG_GAP

    total_image_h = sum(final_cys[sid] for sid, _ in col_images)
    total_needed = total_image_h + total_label_h + total_gaps

    scale = 1.0
    if total_needed > available and total_image_h > 0:
        space_for_images = available - total_label_h - total_gaps
        if space_for_images > 0:
            scale = space_for_images / total_image_h
            logger.info("    [layout] scaling images to %.1f%% to fit slide", scale * 100)
    # Apply scale
    for sid in final_cys:
        final_cys[sid] = int(final_cys[sid] * scale)

    # ── Top-to-bottom layout pass ──
    positions = {}  # sid -> new_y
    sizes = {}      # sid -> new_cy

    # Anchor label stays in place
    if anchor_label:
        cursor = anchor_label[1]["y"] + anchor_label[1]["cy"] + LABEL_IMG_GAP
    else:
        cursor = first_img_y

    for idx, (img_sid, img_geo) in enumerate(col_images):
        # Place image
        positions[img_sid] = cursor
        sizes[img_sid] = final_cys[img_sid]
        cursor += final_cys[img_sid]

        # Place section labels after this image (before next image)
        section_labels = labels_per_section.get(idx, [])
        section_labels.sort(key=lambda x: x[1]["y"])
        if section_labels or idx < len(col_images) - 1:
            cursor += SECTION_GAP
        for lbl_sid, lbl_geo in section_labels:
            positions[lbl_sid] = cursor
            cursor += lbl_geo["cy"] + LABEL_IMG_GAP

    # Position overlays to match their parent image
    for ov_sid, parent_sid in overlay_parent.items():
        if parent_sid in positions:
            positions[ov_sid] = positions[parent_sid]
        if parent_sid in sizes:
            sizes[ov_sid] = sizes[parent_sid]

    # ── STEP 3: Apply ALL changes to XML in one bottom-to-top pass ──
    # Collect all (sid, new_y, new_cy) sorted by current y descending
    all_sids = set(positions.keys()) | set(sizes.keys())
    updates = []
    for sid in all_sids:
        geo = _read_xml_geo(sid)
        cur_y = geo["y"] if geo else 0
        updates.append((sid, cur_y))
    updates.sort(key=lambda x: -x[1])  # bottom-to-top

    off_write_re = re.compile(
        r'(<[^>]*?:off\b[^>]*?\bx\s*=\s*")(\d+)("[^>]*?\by\s*=\s*")(\d+)(")'
    )
    ext_write_re = re.compile(
        r'(<[^>]*?:ext\b[^>]*?\bcx\s*=\s*")(\d+)("[^>]*?\bcy\s*=\s*")(\d+)(")'
    )

    for sid, _ in updates:
        span = _find_shape_span(xml_str, sid)
        if not span:
            continue
        ss, se = span
        shape_xml = xml_str[ss:se]
        modified = False

        # Update y position
        if sid in positions:
            new_y = positions[sid]
            m = off_write_re.search(shape_xml)
            if m and int(m.group(4)) != new_y:
                shape_xml = off_write_re.sub(
                    lambda m_: m_.group(1) + m_.group(2) + m_.group(3) + str(new_y) + m_.group(5),
                    shape_xml, count=1
                )
                modified = True

        # Update cy
        if sid in sizes:
            new_cy = sizes[sid]
            m = ext_write_re.search(shape_xml)
            if m and int(m.group(4)) != new_cy:
                shape_xml = ext_write_re.sub(
                    lambda m_: m_.group(1) + m_.group(2) + m_.group(3) + str(new_cy) + m_.group(5),
                    shape_xml, count=1
                )
                modified = True

        if modified:
            xml_str = xml_str[:ss] + shape_xml + xml_str[se:]
            log_parts = []
            if sid in positions:
                log_parts.append(f"y={positions[sid]}")
            if sid in sizes:
                log_parts.append(f"cy={sizes[sid]}")
            logger.info("    [layout] id=%s: %s", sid, ", ".join(log_parts))

    return xml_str


def _unify_image_heights(images_with_geometry: list, all_shapes: list):
    """Set all images in the same column to the same (maximum) cy.

    Modifies _computed dicts in-place.

    Args:
        images_with_geometry: list of image dicts with _computed and target_shape_id
        all_shapes: list of shape dicts from config (to find geometry)
    """
    if len(images_with_geometry) < 2:
        return

    half_slide = SLIDE_WIDTH_EMU // 2

    # Group images by column
    left_images = []
    right_images = []
    for img in images_with_geometry:
        shape_id = str(img.get("target_shape_id", ""))
        geo = None
        for s in all_shapes:
            if str(s.get("shape_id", "")) == shape_id:
                geo = s.get("geometry")
                break
        if not geo:
            continue
        center_x = geo.get("x", 0) + geo.get("cx", 0) // 2
        if center_x < half_slide:
            left_images.append(img)
        else:
            right_images.append(img)

    # Unify each column
    for group in [left_images, right_images]:
        if len(group) < 2:
            continue
        cy_values = [img["_computed"].get("cy", 0) for img in group]
        max_cy = max(cy_values)
        if max_cy <= 0:
            continue
        for img in group:
            if img["_computed"].get("cy", 0) != max_cy:
                logger.info("    [ok] Uniform sizing: image id=%s cy %d -> %d",
                            img.get("target_shape_id", ""), img["_computed"].get("cy", 0), max_cy)
                img["_computed"]["cy"] = max_cy


# ── Image injection ──────────────────────────────────────────────────────

def _inject_images(raw_dir: str, slide_num: int, images: list):
    """
    For each dynamic image with a resolved_source, copy the source file
    into _raw/ppt/media/ replacing the original media file.
    Validates format compatibility, file size, and image headers before copying.
    """
    # Compatible format groups: target ext -> set of acceptable source exts
    # PNG and JPEG sources can replace ANY image placeholder type
    _FORMAT_COMPAT = {
        ".png": {".png", ".jpg", ".jpeg"},
        ".jpg": {".png", ".jpg", ".jpeg"},
        ".jpeg": {".png", ".jpg", ".jpeg"},
        ".emf": {".png", ".jpg", ".jpeg"},
        ".svg": {".png", ".jpg", ".jpeg"},
        ".wdp": {".png", ".jpg", ".jpeg"},
    }

    # Known image magic byte signatures: (prefix_bytes, label)
    _MAGIC_SIGNATURES = [
        (b"\x89PNG", "PNG"),
        (b"\xff\xd8\xff", "JPEG"),
    ]

    def _format_size(size_bytes: int) -> str:
        """Format byte count to human-readable string."""
        if size_bytes >= 1_048_576:
            return f"{size_bytes / 1_048_576:.1f} MB"
        elif size_bytes >= 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes} B"

    for image in images:
        if not image.get("is_dynamic"):
            continue
        resolved_source = image.get("resolved_source")
        if not resolved_source:
            continue

        # Derive media filename from the relationship target (e.g. "../media/image40.png")
        target = image.get("target", "")
        target_filename = os.path.basename(target) if target else ""
        if not target_filename:
            logger.warning("    [!] Image entry has resolved_source but no target path, skipping")
            continue

        target_path = os.path.join(raw_dir, "ppt", "media", target_filename)
        source_path = resolved_source

        if not os.path.exists(source_path):
            logger.warning("    [!] Image source not found: %s", source_path)
            continue

        # --- Validation 1: Format compatibility ---
        source_ext = os.path.splitext(source_path)[1].lower()
        target_ext = os.path.splitext(target_filename)[1].lower()
        compatible_exts = _FORMAT_COMPAT.get(target_ext)
        # Reject if source format is not in the compat set at all
        format_rejected = (compatible_exts is not None and source_ext not in compatible_exts)
        # Even if compatible, rename when actual extensions differ (JPEG in .png won't render)
        _same_family = {".jpg", ".jpeg"}
        source_norm = source_ext if source_ext not in _same_family else ".jpeg"
        target_norm = target_ext if target_ext not in _same_family else ".jpeg"
        format_mismatch = format_rejected or (source_norm != target_norm)

        # --- Validation 2: File size sanity check ---
        source_size = os.path.getsize(source_path)
        if os.path.exists(target_path):
            target_size = os.path.getsize(target_path)
            if target_size > 0 and source_size > 10 * target_size:
                logger.warning("    [WARNING] Image %s: replacement is %s vs original %s — may affect file size", target_filename, _format_size(source_size), _format_size(target_size))

        # --- Validation 3: Basic header validation ---
        try:
            with open(source_path, "rb") as f:
                header = f.read(8)
            matched_signature = False
            for magic, label in _MAGIC_SIGNATURES:
                if header[:len(magic)] == magic:
                    matched_signature = True
                    break
            if not matched_signature:
                logger.warning("    [WARNING] File may not be a valid image: %s", source_path)
        except Exception:
            logger.warning("    [WARNING] File may not be a valid image: %s", source_path)

        # --- Handle format mismatch: rename target and update rels ---
        if format_mismatch:
            # Map source extension to the OOXML-expected extension
            ext_map = {".jpg": ".jpeg", ".jpeg": ".jpeg", ".png": ".png"}
            new_ext = ext_map.get(source_ext, source_ext)
            new_target_filename = os.path.splitext(target_filename)[0] + new_ext
            new_target_path = os.path.join(raw_dir, "ppt", "media", new_target_filename)

            try:
                shutil.copy2(source_path, new_target_path)
                # Remove old file if different name
                if new_target_filename != target_filename and os.path.exists(target_path):
                    os.remove(target_path)
                # Update the .rels file to point to new filename
                rels_path = os.path.join(raw_dir, "ppt", "slides", "_rels", f"slide{slide_num}.xml.rels")
                if os.path.exists(rels_path):
                    with open(rels_path, "rb") as f:
                        rels_content = f.read().decode("utf-8")
                    old_target_rel = f"../media/{target_filename}"
                    new_target_rel = f"../media/{new_target_filename}"
                    if old_target_rel in rels_content:
                        rels_content = rels_content.replace(old_target_rel, new_target_rel)
                        with open(rels_path, "wb") as f:
                            f.write(rels_content.encode("utf-8"))
                        logger.info("    [ok] Updated rels: %s -> %s", old_target_rel, new_target_rel)
                logger.info("    [ok] Image replaced (format fixed): %s <- %s", new_target_filename, source_path)
            except Exception as e:
                logger.warning("    [!] Failed to copy image %s -> %s: %s", source_path, new_target_path, e)
        else:
            try:
                shutil.copy2(source_path, target_path)
                logger.info("    [ok] Image replaced: %s <- %s", target_filename, source_path)
            except Exception as e:
                logger.warning("    [!] Failed to copy image %s -> %s: %s", source_path, target_path, e)


# ── Token filtering ──────────────────────────────────────────────────────

def _filter_resolved_tokens(resolved_tokens: dict) -> dict:
    """
    Return only tokens whose values are actually resolved.
    - If a token has '_resolved' flag, use that.
    - Otherwise, fall back: skip values that look like unresolved {{...}} refs.
    - Plain string values without {{ }} are considered resolved (backward compat).
    """
    actual = {}
    for k, v in resolved_tokens.items():
        # Check for _resolved flag (set by update_config.py)
        if isinstance(v, dict):
            # Token value might be a dict with 'value' and '_resolved' keys
            # Also handles dicts with 'target_run' — extract the value field
            if v.get("_resolved", True):
                actual[k] = str(v.get("value", ""))
            continue
        # Plain string value — backward compatibility
        v_str = str(v)
        if v_str.startswith("{{") and v_str.endswith("}}"):
            # Looks like an unresolved template reference — skip
            continue
        actual[k] = v_str
    return actual


# ── Main slide injection (string-based) ──────────────────────────────────

def inject_slide(slide_xml_path: str, shapes_to_inject: list, dry_run: bool = False) -> int:
    """
    Opens a slide XML as raw text, injects all token replacements using string
    operations, and writes the modified text back. Never uses tree.write().
    Returns total number of replacements made.
    When dry_run=True, computes and logs changes but skips writing files.
    """
    # Read raw XML preserving everything
    with open(slide_xml_path, "rb") as f:
        raw_bytes = f.read()

    # Detect encoding from XML declaration, default to UTF-8
    xml_str = raw_bytes.decode("utf-8")

    # Preserve the original XML declaration line (including standalone="yes")
    xml_decl = ""
    xml_decl_match = re.match(r'(<\?xml\b[^?]*?\?>)\s*', xml_str)
    if xml_decl_match:
        xml_decl = xml_decl_match.group(1)

    total_replacements = 0

    _cached_root = None
    _xml_dirty = True

    def _get_root():
        nonlocal _cached_root, _xml_dirty
        if _xml_dirty or _cached_root is None:
            _cached_root = _parse_xml_readonly(xml_str)
            _xml_dirty = False
        return _cached_root

    for shape in shapes_to_inject:
        shape_name = shape.get("shape_name", "")
        shape_id = str(shape.get("shape_id", ""))

        # Route table shapes to dedicated table injection
        if shape.get("category") == "table":
            resolved_value = shape.get("resolved_value")
            resolved_tokens = shape.get("resolved_tokens", {})
            if resolved_value or resolved_tokens:
                xml_str, table_count = _inject_table_data(xml_str, shape)
                total_replacements += table_count
                _xml_dirty = True
            continue

        # ── New path: direct field replacement via resolved_value ──
        resolved_value = shape.get("resolved_value")
        if resolved_value is not None:
            # Use ET for read-only query to verify shape exists
            root = _get_root()
            cnvpr = _find_shape_by_id(root, shape_id)
            if cnvpr is None:
                logger.warning("    [!] Shape not found in XML by id=%s: '%s'", shape_id, shape_name)
                continue

            shape_node = _get_shape_element(cnvpr, root)
            if shape_node is None:
                logger.warning("    [!] Could not locate shape element for id=%s: '%s'", shape_id, shape_name)
                continue

            original_len = _get_shape_text_length(shape_node)

            span = _find_shape_span(xml_str, shape_id)
            if span is None:
                logger.warning("    [!] Could not locate shape span in raw XML for id=%s: '%s'", shape_id, shape_name)
                continue

            start, end = span
            shape_xml = xml_str[start:end]

            # target_run: replace text in a specific run only, preserving all others
            target_run = shape.get("target_run")
            if target_run is not None and isinstance(target_run, int):
                at_pattern_tr = re.compile(
                    r'(<[^>]*?:t\b[^>]*?>)(.*?)(</[^>]*?:t\s*>)',
                    re.DOTALL
                )
                at_matches_tr = list(at_pattern_tr.finditer(shape_xml))
                if at_matches_tr and 0 <= target_run < len(at_matches_tr):
                    new_shape_xml = shape_xml
                    am = at_matches_tr[target_run]
                    # Preserve leading/trailing whitespace from the original run text
                    orig_text = am.group(2)
                    orig_unescaped = orig_text.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'")
                    leading_ws = orig_unescaped[:len(orig_unescaped) - len(orig_unescaped.lstrip())]
                    trailing_ws = orig_unescaped[len(orig_unescaped.rstrip()):]
                    escaped_val = _escape_for_xml(leading_ws + resolved_value + trailing_ws)
                    new_at = am.group(1) + escaped_val + am.group(3)
                    new_shape_xml = new_shape_xml[:am.start()] + new_at + new_shape_xml[am.end():]
                    xml_str = xml_str[:start] + new_shape_xml + xml_str[end:]
                    _xml_dirty = True
                    total_replacements += 1
                    logger.info("    [ok] %s (id=%s): replaced [target_run=%d]", shape_name, shape_id, target_run)
                    logger.info("      -> '%s'", resolved_value[:80])

                    # Font sizing after target_run replacement
                    layout = shape.get("layout", {})
                    computed = layout.get("_computed")
                    max_font_size = layout.get("max_font_size")
                    if computed and computed.get("font_size") and max_font_size and computed["font_size"] < max_font_size:
                        target_sz = computed["font_size"]
                        new_span = _find_shape_span(xml_str, shape_id)
                        if new_span is not None:
                            s, e = new_span
                            fitted = _apply_text_autofit(xml_str[s:e], target_sz)
                            xml_str = xml_str[:s] + fitted + xml_str[e:]
                            _xml_dirty = True
                        logger.info("    [ok] %s (id=%s): target_run auto-fit font=%d", shape_name, shape_id, target_sz)
                    elif len(resolved_value) > len(am.group(2)):
                        geo = shape.get("geometry", {})
                        shape_cx = geo.get("cx", 0)
                        shape_cy = geo.get("cy", 0)
                        max_sz = layout.get("max_font_size", 3200)
                        min_sz = layout.get("min_font_size", 400)
                        if shape_cx > 0 and shape_cy > 0:
                            from layout import compute_text_font_scale
                            target_sz = compute_text_font_scale(
                                resolved_value, shape_cx, shape_cy,
                                max_sz, min_sz, max_sz
                            )
                            if target_sz < max_sz:
                                new_span = _find_shape_span(xml_str, shape_id)
                                if new_span is not None:
                                    s, e = new_span
                                    fitted = _apply_text_autofit(xml_str[s:e], target_sz)
                                    xml_str = xml_str[:s] + fitted + xml_str[e:]
                                    _xml_dirty = True
                                logger.info("    [ok] %s (id=%s): target_run computed font=%d", shape_name, shape_id, target_sz)
                    continue
                else:
                    logger.warning("    [!] %s (id=%s): target_run=%d out of range (%d runs), falling back",
                                   shape_name, shape_id, target_run, len(at_matches_tr))

            new_shape_xml, modified = _replace_shape_text(shape_xml, resolved_value)

            if modified:
                # Expand any newlines in <a:t> into separate <a:p> paragraphs
                new_shape_xml = _expand_newlines_to_paragraphs(new_shape_xml)
                xml_str = xml_str[:start] + new_shape_xml + xml_str[end:]
                _xml_dirty = True
                total_replacements += 1

                # Font sizing — prefer layout-computed auto-fit, fall back to proportional shrink
                layout = shape.get("layout", {})
                computed = layout.get("_computed")
                new_len = len(resolved_value)

                max_font_size = layout.get("max_font_size") or shape.get("max_font_size")

                if computed and computed.get("font_size") and max_font_size and computed["font_size"] < max_font_size:
                    # Overflow detected: pre-computed font size is smaller than max — apply auto-fit
                    target_sz = computed["font_size"]
                    new_span = _find_shape_span(xml_str, shape_id)
                    if new_span is not None:
                        s, e = new_span
                        fitted = _apply_text_autofit(xml_str[s:e], target_sz)
                        xml_str = xml_str[:s] + fitted + xml_str[e:]
                        _xml_dirty = True
                    logger.info("    [ok] %s (id=%s): replaced [auto-fit font=%d]", shape_name, shape_id, target_sz)
                elif new_len > original_len:
                    # Compute proper font that fits using shape geometry
                    geo = shape.get("geometry", {})
                    shape_cx = geo.get("cx", 0)
                    shape_cy = geo.get("cy", 0)
                    max_sz = layout.get("max_font_size", 3200)
                    min_sz = layout.get("min_font_size", 400)
                    if shape_cx > 0 and shape_cy > 0:
                        from layout import compute_text_font_scale
                        target_sz = compute_text_font_scale(
                            resolved_value, shape_cx, shape_cy,
                            max_sz, min_sz, max_sz
                        )
                        if target_sz < max_sz:
                            new_span = _find_shape_span(xml_str, shape_id)
                            if new_span is not None:
                                s, e = new_span
                                fitted = _apply_text_autofit(xml_str[s:e], target_sz)
                                xml_str = xml_str[:s] + fitted + xml_str[e:]
                                _xml_dirty = True
                            logger.info("    [ok] %s (id=%s): replaced [auto-fit font=%d]", shape_name, shape_id, target_sz)
                        else:
                            logger.info("    [ok] %s (id=%s): replaced [fits at original font]", shape_name, shape_id)
                    else:
                        # No geometry — fall back to proportional shrink
                        new_span = _find_shape_span(xml_str, shape_id)
                        if new_span is not None:
                            s, e = new_span
                            shrunk = _shrink_font_in_shape_xml(xml_str[s:e], original_len, new_len)
                            xml_str = xml_str[:s] + shrunk + xml_str[e:]
                            _xml_dirty = True
                        logger.info("    [ok] %s (id=%s): replaced [font scaled %s->%s chars]", shape_name, shape_id, original_len, new_len)
                else:
                    logger.info("    [ok] %s (id=%s): replaced", shape_name, shape_id)

                logger.info("      -> '%s'", resolved_value[:80])
            else:
                logger.warning("    [!] %s (id=%s): no <a:t> elements found in shape", shape_name, shape_id)
            continue

        # ── Legacy path: token-based replacement via resolved_tokens ──
        resolved_tokens = shape.get("resolved_tokens", {})
        if not resolved_tokens:
            continue

        actual_tokens = _filter_resolved_tokens(resolved_tokens)
        if not actual_tokens:
            logger.info("    [skip] %s (id=%s) -- no resolved values", shape_name, shape_id)
            continue

        root = _get_root()
        cnvpr = _find_shape_by_id(root, shape_id)
        if cnvpr is None:
            logger.warning("    [!] Shape not found in XML by id=%s: '%s'", shape_id, shape_name)
            continue

        shape_node = _get_shape_element(cnvpr, root)
        if shape_node is None:
            logger.warning("    [!] Could not locate shape element for id=%s: '%s'", shape_id, shape_name)
            continue

        original_len = _get_shape_text_length(shape_node)

        span = _find_shape_span(xml_str, shape_id)
        if span is None:
            logger.warning("    [!] Could not locate shape span in raw XML for id=%s: '%s'", shape_id, shape_name)
            continue

        start, end = span
        shape_xml = xml_str[start:end]

        shape_count = 0
        for token, new_value in actual_tokens.items():
            modified_shape_xml, count = _replace_tokens_in_shape_xml(shape_xml, {token: new_value})
            shape_xml = modified_shape_xml
            shape_count += count

            if count > 0:
                xml_str = xml_str[:start] + shape_xml + xml_str[end:]
                _xml_dirty = True
                end = start + len(shape_xml)

                root = _get_root()
                cnvpr = _find_shape_by_id(root, shape_id)
                if cnvpr is not None:
                    shape_node = _get_shape_element(cnvpr, root)

                new_span = _find_shape_span(xml_str, shape_id)
                if new_span is not None:
                    start, end = new_span
                    shape_xml = xml_str[start:end]

        total_replacements += shape_count

        if shape_count > 0:
            # Token replacement is surgical — never shrink fonts.
            # The user is swapping content tokens, not rewriting the whole shape.
            logger.info("    [ok] %s (id=%s): %s replacement(s)", shape_name, shape_id, shape_count)

            for token, value in actual_tokens.items():
                logger.info("      '%s' -> '%s'", token, value)
        else:
            logger.warning("    [!] %s (id=%s): tokens defined but none matched in paragraphs", shape_name, shape_id)

    # Ensure the original XML declaration is preserved
    if xml_decl:
        decl_check = re.match(r'(<\?xml\b[^?]*?\?>)\s*', xml_str)
        if decl_check:
            xml_str = xml_decl + xml_str[decl_check.end(1):]
        else:
            xml_str = xml_decl + "\n" + xml_str

    # Write back as raw bytes — NEVER use tree.write()
    if not dry_run:
        with open(slide_xml_path, "wb") as f:
            f.write(xml_str.encode("utf-8"))

    return total_replacements


def _update_content_types(raw_dir: str):
    """Update [Content_Types].xml if image extensions changed during injection."""
    ct_path = os.path.join(raw_dir, "[Content_Types].xml")
    if not os.path.exists(ct_path):
        return

    with open(ct_path, "r", encoding="utf-8") as f:
        ct_xml = f.read()

    # Scan actual media files present in _raw/ppt/media/
    media_dir = os.path.join(raw_dir, "ppt", "media")
    if not os.path.exists(media_dir):
        return

    # Map of extension -> content type
    EXT_TO_CONTENT_TYPE = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".emf": "image/x-emf",
        ".wmf": "image/x-wmf",
        ".svg": "image/svg+xml",
        ".tiff": "image/tiff",
        ".tif": "image/tiff",
        ".wdp": "image/vnd.ms-photo",
    }

    # Collect all extensions actually used in media/
    used_exts = set()
    for fname in os.listdir(media_dir):
        ext = os.path.splitext(fname)[1].lower()
        if ext:
            used_exts.add(ext)

    # For each used extension, ensure a Default entry exists in [Content_Types].xml
    modified = False
    for ext in used_exts:
        ext_no_dot = ext.lstrip(".")
        content_type = EXT_TO_CONTENT_TYPE.get(ext)
        if not content_type:
            continue
        # Check if a Default for this extension already exists
        pattern = re.compile(r'<Default\s+[^>]*Extension\s*=\s*"' + re.escape(ext_no_dot) + r'"', re.IGNORECASE)
        if not pattern.search(ct_xml):
            # Insert before </Types>
            insert_tag = f'<Default Extension="{ext_no_dot}" ContentType="{content_type}"/>'
            ct_xml = ct_xml.replace("</Types>", insert_tag + "\n</Types>")
            modified = True

    if modified:
        with open(ct_path, "w", encoding="utf-8") as f:
            f.write(ct_xml)


def inject(config_path: str = "configs/slides_examples.json", library_path: str = "component_library", dry_run: bool = False):
    logger.info("Injecting from config: %s", config_path)
    logger.info("Library: %s", library_path)

    if dry_run:
        logger.info("[DRY RUN] No files will be modified")

    if not os.path.exists(config_path):
        logger.error("[ERROR] Config not found: %s", config_path)
        sys.exit(1)

    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)

    raw_dir = os.path.join(library_path, "_raw")
    if not os.path.exists(raw_dir):
        logger.error("[ERROR] _raw/ not found at %s. Run deconstruct.py first.", raw_dir)
        sys.exit(1)

    # Ensure idempotency: always inject from a pristine _raw/ state (skip in dry-run mode)
    clean_dir = os.path.join(library_path, "_raw_clean")
    if not dry_run:
        if not os.path.exists(clean_dir):
            # First run: save the pristine _raw/ as _raw_clean
            shutil.copytree(raw_dir, clean_dir)
            logger.info("  Clean backup created: %s", clean_dir)
        else:
            # Subsequent runs: restore _raw/ from _raw_clean so we inject from a clean state
            shutil.rmtree(raw_dir)
            shutil.copytree(clean_dir, raw_dir)
            logger.info("  Restored _raw/ from clean backup: %s", clean_dir)

    total_replacements = 0
    slides_modified = 0

    try:
        for slide_key, slide in config["slides"].items():
            slide_num = slide["slide_number"]

            # ── Text shapes ──────────────────────────────────────────────
            # Resolve literal: fields inline before filtering
            for s in slide["shapes"]:
                df = s.get("data_field", "")
                if df.startswith("literal:") and s.get("is_dynamic"):
                    s["resolved_value"] = df[len("literal:"):]

            dynamic_shapes = [
                s for s in slide["shapes"]
                if s.get("is_dynamic") and (s.get("resolved_value") is not None or s.get("resolved_tokens"))
            ]

            # ── Dynamic images ───────────────────────────────────────────
            dynamic_images = [
                img for img in slide.get("images", [])
                if img.get("is_dynamic") and img.get("resolved_source")
            ]

            if not dynamic_shapes and not dynamic_images:
                continue

            slide_xml_path = os.path.join(raw_dir, "ppt", "slides", f"slide{slide_num}.xml")

            logger.info("\n  [ Slide %s ]", slide_num)

            # Inject text tokens
            if dynamic_shapes:
                if not os.path.exists(slide_xml_path):
                    logger.warning("  [!] Slide XML not found: %s", slide_xml_path)
                else:
                    count = inject_slide(slide_xml_path, dynamic_shapes, dry_run=dry_run)
                    total_replacements += count
                    if count > 0:
                        slides_modified += 1

            # Inject images (skip in dry-run mode)
            if dynamic_images and not dry_run:
                _inject_images(raw_dir, slide_num, dynamic_images)

                # Apply image geometry (resize pic shapes in slide XML)
                images_with_geometry = [
                    img for img in dynamic_images
                    if img.get("_computed") and img.get("target_shape_id")
                ]
                if images_with_geometry and os.path.exists(slide_xml_path):
                    with open(slide_xml_path, "rb") as f:
                        img_xml_str = f.read().decode("utf-8")

                    # Clear srcRect on dynamic images (reset cropping)
                    for img in images_with_geometry:
                        shape_id = str(img["target_shape_id"])
                        span = _find_shape_span(img_xml_str, shape_id)
                        if span:
                            s_start, s_end = span
                            shape_xml = img_xml_str[s_start:s_end]
                            srcrect_pat = re.compile(r'<[^>]*?:srcRect\b[^>]*/>')
                            srcrect_full = re.compile(r'<[^>]*?:srcRect\b[^>]*>.*?</[^>]*?:srcRect\s*>', re.DOTALL)
                            if srcrect_full.search(shape_xml):
                                shape_xml = srcrect_full.sub('<a:srcRect l="0" t="0" r="0" b="0"/>', shape_xml, count=1)
                            elif srcrect_pat.search(shape_xml):
                                shape_xml = srcrect_pat.sub('<a:srcRect l="0" t="0" r="0" b="0"/>', shape_xml, count=1)
                            img_xml_str = img_xml_str[:s_start] + shape_xml + img_xml_str[s_end:]

                    # Single consolidated layout pass — computes final cy for
                    # every image, lays out labels + images top-to-bottom,
                    # and applies all y/cy changes to XML in one pass.
                    img_xml_str = _layout_image_sections(
                        img_xml_str, images_with_geometry, slide["shapes"]
                    )

                    with open(slide_xml_path, "wb") as f:
                        f.write(img_xml_str.encode("utf-8"))

        # Update content types if image formats changed (skip in dry-run mode)
        if not dry_run:
            _update_content_types(raw_dir)

    except Exception as e:
        if not dry_run:
            logger.error("\n[ERROR] Injection failed: %s", e)
            logger.info("  Rolling back _raw/ from clean backup: %s", clean_dir)
            shutil.rmtree(raw_dir)
            shutil.copytree(clean_dir, raw_dir)
            logger.info("  Rollback complete. _raw/ restored to pre-injection state.")
        raise

    if dry_run:
        logger.info("[DRY RUN] Would have modified %d slides with %d replacements", slides_modified, total_replacements)
    else:
        logger.info("\n[DONE] Injection complete.")
        logger.info("   Slides modified:    %s", slides_modified)
        logger.info("   Total replacements: %s", total_replacements)


if __name__ == "__main__":
    config = sys.argv[1] if len(sys.argv) > 1 else "configs/slides_examples.json"
    lib = sys.argv[2] if len(sys.argv) > 2 else "component_library"
    inject(config, lib)
