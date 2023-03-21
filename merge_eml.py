import mailbox
from email.parser import BytesParser
from pathlib import Path
import sys

# This script merges EML files into a single mbox file. Run pip/pip3 install mailbox before running it, then run it as follows:
# python/python3 merge_eml.py <eml_folder> <mbox_output_file>

def main():
    eml_folder = sys.argv[1]  # folder containing your EML files
    mbox_file = sys.argv[2]  # Output mbox file name

    mbox = mailbox.mbox(mbox_file)
    mbox.lock()

    try:
        for eml_file in Path(eml_folder).glob("*.eml"):
            with open(eml_file, "rb") as file:
                msg = BytesParser().parse(file)
                mbox.add(msg)
    finally:
        mbox.unlock()
        mbox.close()

    print(f"Merged EML files into {mbox_file}")

if __name__ == "__main__":
  
    if len(sys.argv) != 3:
        arg_number = len(sys.argv)
        arg_number = arg_number - 1
        arg_number = str(arg_number)
        
        print ("Invalid number of arguments. You've passed " + arg_number + " arguments, but 2 are required.")
        print("Usage: python/python3 merge_eml.py <eml_folder> <mbox_output_file>")
        sys.exit(1)
    print("EMLs folder: "+ sys.argv[1])
    print("MBOX output file: merged.mbox")
    main()
