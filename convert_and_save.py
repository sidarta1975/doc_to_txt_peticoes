import os
import docx2txt

def convert_and_save(input_folder, output_folder):
    """
    Converts.doc and.docx files to.txt and saves them in a specified output folder.

    Args:
        input_folder (str): Path to the folder containing the input files.
        output_folder (str): Path to the folder where the.txt files should be saved.
    """

    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.endswith((".doc", ".docx")):
                input_file_path = os.path.join(root, file)
                output_file_path = os.path.join(output_folder, f"{os.path.splitext(file)}.txt")

                try:
                    text = docx2txt.process(input_file_path)
                    with open(output_file_path, "w", encoding="utf-8") as f:
                        f.write(text)
                    print(f"Converted '{input_file_path}' to '{output_file_path}'")
                except Exception as e:
                    print(f"Error converting '{input_file_path}': {e}")

if __name__ == "__main__":
    input_folder = "/Users/sbm/Library/CloudStorage/GoogleDrive-sidarta.martins@gmail.com/.shortcut-targets-by-id/1YAj25RdMk2FNbI73sxowN3ImxceK7Dh_/Advogada Parceira/Petições/peticoes_doc_docx/modelos_peticao/Previdenciario"  # Replace with the actual path
    output_folder = "/Users/sbm/Library/CloudStorage/GoogleDrive-sidarta.martins@gmail.com/.shortcut-targets-by-id/1YAj25RdMk2FNbI73sxowN3ImxceK7Dh_/Advogada Parceira/Petições/peticoes_txt/previdenciario"  # Replace with the actual path

    convert_and_save(input_folder, output_folder)