import sys, os

from docxtemplater.templater import write

if __name__ == '__main__':
    if len(sys.argv) == 4:
        template = sys.argv[1]
        replacements = sys.argv[2]
        output = sys.argv[3]
        if not os.path.exists(template):
            print("Template file not found", file=sys.stderr)
        write(template, replacements, output)
    else:
        print("Usage: docxtemplater.py template replacements, output", file=sys.stderr)
