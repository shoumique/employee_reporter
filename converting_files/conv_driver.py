import argparse
import os
from typing import Any

import pandas as pd
import unicodeconverter as uc


def _looks_like_bijoy(text: str) -> bool:
	"""Heuristic: return True if the text is likely Bijoy (non‑Unicode Bangla).

	We assume Bijoy text is mostly ASCII/extended Latin and *does not* contain
	characters from the Bengali Unicode block (U+0980–U+09FF). Pure numbers/
	punctuation are treated as non-Bijoy.
	"""

	if not text or text.isspace():
		return False

	has_alpha = False
	for ch in text:
		code = ord(ch)
		# If we see any Bengali-range character, treat as already-Unicode Bangla
		if 0x0980 <= code <= 0x09FF:
			return False
		if ch.isalpha():
			has_alpha = True

	# No Bengali letters, but has some alphabetic characters -> probably Bijoy
	return has_alpha


def convert_bijoy_in_value(value: Any) -> Any:
	"""Convert a value from Bijoy to Unicode only when it looks like Bijoy text."""

	if isinstance(value, str) and value.strip() and _looks_like_bijoy(value):
		try:
			return uc.convert_bijoy_to_unicode(value)
		except Exception:
			# If anything goes wrong, fall back to original value
			return value
	return value


def convert_excel_bijoy_to_unicode(input_path: str, output_csv_path: str) -> None:
	"""Read an Excel file, convert Bijoy-encoded Bangla text to Unicode, and save as CSV."""

	# Read all sheets and concatenate, or just the first? For generality, use first sheet.
	df = pd.read_excel(input_path)

	# Apply conversion to every cell
	df_converted = df.applymap(convert_bijoy_in_value)

	# Also convert column headers and index labels, which pandas keeps separate
	df_converted.columns = [convert_bijoy_in_value(col) for col in df_converted.columns]
	df_converted.index = [convert_bijoy_in_value(idx) for idx in df_converted.index]

	# Ensure output directory exists
	os.makedirs(os.path.dirname(output_csv_path) or ".", exist_ok=True)

	df_converted.to_csv(output_csv_path, index=False)


def build_default_output_path(input_path: str) -> str:
	base, _ = os.path.splitext(input_path)
	return f"{base}_unicode.csv"


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description="Convert Bijoy Bangla text in an Excel file to Unicode and export as CSV.",
	)
	parser.add_argument(
		"input_excel",
		help="Path to the input Excel file (e.g. bank sheet).",
	)
	parser.add_argument(
		"-o",
		"--output-csv",
		help="Path to output CSV file. Default: <input_basename>_unicode.csv",
	)
	return parser.parse_args()


def main() -> None:
	args = parse_args()
	input_path = args.input_excel

	if not os.path.isfile(input_path):
		raise FileNotFoundError(f"Input Excel file not found: {input_path}")

	output_path = args.output_csv or build_default_output_path(input_path)

	convert_excel_bijoy_to_unicode(input_path, output_path)

	print(f"Converted file written to: {output_path}")


if __name__ == "__main__":
	main()

