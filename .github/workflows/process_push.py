# Contact @danguetta with any questions

import os
import shutil
from oletools.olevba3 import VBA_Parser

EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)

def parse(workbook_path, extract_path, KEEP_NAME=False):
    '''
    Given the path of a workbook, this will extract the VBA code from the workbook
    in workbook_path to the folder in extract_path.
    
    If KEEP_NAME is True, we keep the line "Attribute VB_Name" in the files
    
    Slightly modified from the code in
       https://www.xltrail.com/blog/auto-export-vba-commit-hook
    '''
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        try:
            decoded_content = content.decode('latin-1')
        except:
            decoded_content = content
            
        lines = []
        if '\r\n' in decoded_content:
            lines = decoded_content.split('\r\n')
        else:
            lines = decoded_content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(extract_path)):
                        os.makedirs(extract_path)
                    with open(os.path.join(extract_path, filename), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content).lower())


if __name__ == '__main__':
    # Define the folder for the VBA code; this folder will be deleted,
    # and the VBA from the Excel will be extracted into this folder
    VBA_CODE_FOLDER = './~VBA Code'

    # Define the expected location and names of the files
    expected_files = {'add-in'       : './CBS BA Multiplatform add-in.xlam',
                      'manual_pdf'   : './User manual/BA Add-In User Manual.pdf',
                      'manual_latex' : './User manual/BA Add-In User Manual.tex'}
    
    # Create a list to store errors
    errors = []
    
    # ---------------------------------------
    # -  Ensure expected files are present  -
    # ---------------------------------------
    
    files_exist = {f : os.path.isfile(expected_files[f]) for f in expected_files}
    
    # If any files are missing, add to the error list
    if not all(files_exist.values()):
        errors.append('  - Some files I expected to find were not present:')
        errors.extend([f'    *  {f} expected at {expected_files[f]} but not found.'
                                        for f in expected_files if not files_exist[f]])
                
    # ---------------------------------------
    # -  Ensure the PDF was generated from  -
    # -  the latest LaTeX file              -
    # ---------------------------------------
    
    if files_exist['manual_pdf'] and files_exist['manual_latex']:
        # Generate the MD5 hash of the LaTeX file
        import hashlib
        with open(expected_files['manual_latex'], 'rb') as f:
            f_text = f.read()
            latex_hash = hashlib.md5(f_text).hexdigest().upper()
            latex_hash_crlf = hashlib.md5(f_text.replace(b'\n', b'\r\n')).hexdigest().upper()
        
        # Load the hash printed to the PDF
        import PyPDF2
        pdf_page = PyPDF2.PdfReader(expected_files['manual_pdf']).pages[0].extract_text()
        pdf_hash = pdf_page.split('Manual version: ')[1].split('\n')[0]
        
        # Ensure they match
        print(f'Latex hash:           {latex_hash}')
        print(f'Latex hash with crlf: {latex_hash_crlf}')
        print(f'Hash from pdf:        {pdf_hash}')
        if pdf_hash not in [latex_hash, latex_hash_crlf]:
            errors.append('  - The user manual PDF file does not seem to have been generated '
                              'the current user manual LaTeX file. Please recompile the latest '
                              'LaTeX file, so that it correctly reflects the PDF.')
        
    # ----------------------
    # -  Extract VBA code  -
    # ----------------------
    
    code_extract_done = False
    if files_exist['add-in']:
        try:
            # If the VBA_CODE_FOLDER does exist, create it
            if not os.path.exists(VBA_CODE_FOLDER):
                os.makedirs(VBA_CODE_FOLDER)
            else:
                # Delete the folder containing the VBA code
                shutil.rmtree(VBA_CODE_FOLDER)
            
            # Extract the VBA code from the workbook
            parse('CBS BA Multiplatform add-in.xlam', VBA_CODE_FOLDER)
            
            code_extract_done = True
        except Exception as e:
            errors.append(f'  - Error extracting the VBA code from the workbook; the error was {str(e)}')
    
    # -------------------------------
    # -  Deal with version numbers  -
    # -------------------------------

    # Get the version number from the VBA file
    vba_version = None
    if code_extract_done:
        try:
            # Load utils.bas, where the version number resides
            with open(VBA_CODE_FOLDER + '/Util') as f:
                lines = f.readlines()
            
            # Find the line that contains the version number
            lines = [i for i in lines if "Const constVersionNumber".lower() in i]
            
            if len(lines) == 0:
                errors.append('  - The Util.bas VBA file does not contain the version number.')
            elif len(lines) > 1:
                errors.append('  - The version number is defined multiple times in Utils.bas')
            else:
                # Extract the version number from the defining line in VBA. The line looks like this
                #     Public Const constVersionNumber As String = "0.0.12"\n
                # We extract it by first splitting on the equal sign, removing leading and trailing
                # whitespace, and then removing the quotes
                vba_version = lines[0].split('=')[1].strip()[1:-1]        
        except Exception as e:
            errors.append(f'  - Error extracting the version number from the VBA files; the error was {str(e)}')

    if vba_version is None:
        errors.append('  - Could not extract the version number from the VBA files')
    
    # Get the version number from the manual
    manual_version = None
    if files_exist['manual_latex']:
        try:
            # Load the manual LaTeX
            with open(expected_files['manual_latex']) as f:
                lines = f.readlines()
            
            # Find the line that contains the version number
            lines = [i for i in lines if '\\newcommand{\\versionNumber}{' in i]
            
            if len(lines) == 0:
                errors.append('  - The user manual latex file does not contain the version number.')
            elif len(lines) > 1:
                errors.append('  - The version number is defined multiple times in the user manual latex file.')
            else:
                # Extract the version number from the line. The line looks like this
                #     \newcommand{\versionNumber}{4.0}
                manual_version = lines[0].split('{')[-1].strip()[:-1]
        except Exception as e:
            errors.append(f'  - Error extracting the version number from the user manual latex file; the error was {str(e)}')
    
    if manual_version is None:
        errors.append('  - Could not extract the version number from the user manual LaTeX file') 

    # Ensure the version numbers match
    if ( (vba_version is not None) and (manual_version is not None)
                                and (vba_version != manual_version) ):
        errors.append(f'  - The version numbers in the VBA file ({vba_version}) and in the user manual '
                                                    f'LaTeX file ({manual_version}) do not match. Please fix.')

    # ----------------------------
    # -  Update the readme file  -
    # ----------------------------
    
    # Read the readme file
    with open('./README.md') as f:
        readme_file = f.readlines()
    
    # Find the line with the comment
    comment_line_n = [i for i, j in enumerate(readme_file) if '<!-- DO ***NOT*** EDIT ANYTHING ABOVE THIS LINE, INCLUDING THIS COMMENT -->' in j]
    assert len(comment_line_n) == 1
    comment_line_n = comment_line_n[0]
    
    # Remove everything before that line
    readme_file = readme_file[comment_line_n:]
    
    # Insert the version number
    if vba_version is not None:
        readme_file.insert(0, f'## Version: {vba_version}\n')
    else:
        readme_file.insert(0, '## Could not find version number in VBA\n')
        
    # If there are errors, insert them into the file
    if len(errors) > 0:
        errors[0:0] = ['## ERROR REPORT',
                       '**Some errors were found while processing your '
                                'latest push. Please fix them before proceeding as follows**',
                       '  - Pull the latest commit (made by the VBA robot) from github',
                       '  - Make your changes',
                       '  - Push back to github',
                       '',
                       'The specific errors I found were as follows:']
        
        errors = [i + '\n' for i in errors]
        
        readme_file[0:0] = errors
    
    # Insert the readme header
    readme_file.insert(0, '# CBS VBA Business Analytics add-in\n')
    
    # Write the file back
    with open ('./README.md', 'w') as f:
        f.writelines(readme_file)