import oletools.olevba as olevba

# Path to your .xlsm file
xlsm_file_path = 'VBA.Calculator_v1.xlsm'

# Open the .xlsm file and extract VBA macros
vba_parser = olevba.VBA_Parser(xlsm_file_path)
output_lines = []

if vba_parser.detect_vba_macros():
    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        output_lines.append(f"Filename: {filename}")
        output_lines.append(f"Stream Path: {stream_path}")
        output_lines.append(f"VBA Filename: {vba_filename}")
        output_lines.append("VBA Code:")
        output_lines.append(vba_code)
        output_lines.append("\n" + "="*80 + "\n")
else:
    output_lines.append("No VBA macros found.")


# Write the output to a text file 
with open('output.txt', 'w') as file: 
    file.write('\n'.join(output_lines))
