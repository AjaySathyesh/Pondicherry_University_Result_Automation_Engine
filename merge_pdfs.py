from PyPDF2 import PdfMerger
import os

class PDFMerger:
    def merge_documents(self, directory, output_path):
        merger = PdfMerger()

        try:
            # Add each PDF file in the directory to the merger
            for filename in sorted(os.listdir(directory)):
                if filename.endswith('.pdf'):
                    file_path = os.path.join(directory, filename)
                    merger.append(file_path)

            # Write the merged PDF to the output path
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            merger.close()
        except Exception as e:
            raise Exception(f"Error merging PDFs: {e}")
