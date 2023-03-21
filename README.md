# merge_eml_mbox
Merges several .elm files into a single .mbox file

The 'convert_olm_to_eml.py' script came from 'https://gist.github.com/tennantje/f04f54871db5e7a0507331f69648f7f4'. Run it as needed and use the 'merge_eml.py' as:

python/python3 merge_eml.py <eml_folder> <mbox_output_file>

for example, "python3 merge_eml.py ./Accounts/<username>@<domain>com.microsoft.__Messages/INBOX merged.mbox"

Afer that, you can import the mbox file into outlook by converting it using Thunderbird, for example.

