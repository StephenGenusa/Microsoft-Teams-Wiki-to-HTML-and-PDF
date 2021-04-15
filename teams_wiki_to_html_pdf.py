#
# Create HTML and PDF files from Microsoft Wiki files
# See github repo for details: https://github.com/StephenGenusa/Microsoft-Teams-Wiki-to-HTML-and-PDF
# By Stephen Genusa, April 2020
#
# Uses Python >= 3.x with no additional requirements
#

import argparse
import os
import re
from subprocess import Popen
from subprocess import PIPE


def strip_pointless_mht_msincomps(mht_file_name):
    # It becomes suggestive, after digging in the internals, why Microsoft hasn't made a PDF or Word export
    # for Microsoft Teams. The wiki editor hides terrors that are common to many of the Microsoft HTML
    # editors I have seen over the years
    with open(mht_file_name, "r", encoding="utf-8") as input_file:
        mht_text = input_file.read()
    html_file_name = os.path.splitext(mht_file_name)[0] + ".html"
    file_path, file_name = os.path.split(html_file_name)
    html_file_name = os.path.join(
        file_path, re.sub(r"(.*)? - (\d{1,4})", r"\2 - \1", file_name)
    )
    print("Writing HTML", html_file_name)
    with open(html_file_name, "w", encoding="utf-8") as output_file:
        # Standardize the font and font-size because the Microsoft editor hides the insane nested font span
        # elements it creates in the HTML
        output_file.write(
            '<html><meta charset="utf-8"><body style="font-size: 18px; font-family: Arial,'
            ' Helvetica, sans-serif;">'
        )
        # throw out the mht header and retain the HTML
        html_text = mht_text[mht_text.find("<h1 id") :]
        # parse the dubious and useless image references with width and height, retaining those
        html_text = re.sub(
            r'<img.*/(img.*)" src.*(height=".*?").*(width=".*?")(/>|>)',
            r'<img src="\1" \2 \3/>',
            html_text,
        )
        # handle all other img cases
        html_text = re.sub(
            r"<img.*/(img-\d{1,4}-.*?\.\w{3,4}).*?>", r'<img src="\1"/>', html_text
        )
        # Oh boy. Matryoshka style font family and size references...
        for loop_count in range(1, 5):
            html_text = re.sub(
                r'<span style="font-fam.*?">(.*)?</span>', r"\1", html_text
            )
            html_text = re.sub(
                r'<span style="font-siz.*?">(.*)?</span>', r"\1", html_text
            )
        output_file.write(html_text)
        return html_file_name


def run_wkhtmltopdf(html_file_name):
    pdf_file_name = os.path.splitext(html_file_name)[0] + ".pdf"
    print("Writing PDF", pdf_file_name)
    command = (
        r'"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" --enable-local-file-access '
        '--page-size Letter "%s" "%s" >> wkhtmlpdf.log'
        % (html_file_name, pdf_file_name)
    )
    try:
        p = Popen(command, shell=True, stdout=PIPE, stderr=PIPE, close_fds=True)
        p.communicate()
    except OSError as exc:
        raise exc


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="Microsoft Teams Wiki to HTML and PDF by Stephen Genusa"
    )
    parser.add_argument(
        "--wiki-input-path",
        help="path to wiki mht files",
        default=os.path.join(os.getenv("USERPROFILE"), "Desktop", "General"),
    )
    args = parser.parse_args()

    for root, dirs, files in os.walk(args.wiki_input_path):
        for file in files:
            if os.path.splitext(file)[1].lower() == ".mht":
                run_wkhtmltopdf(strip_pointless_mht_msincomps(os.path.join(root, file)))
