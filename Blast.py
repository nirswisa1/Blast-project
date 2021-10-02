from Bio.Blast import NCBIWWW
import xml.etree.ElementTree as ET
import xlsxwriter

base = r"C:\\Users\\nir\\Desktop\\nir\\"
input_file = base + "input.fasta"

sequence_data = open(input_file).read() 

result_handle = NCBIWWW.qblast("blastp", "nr", sequence_data, entrez_query='"Nematostella vectensis"[organism]', format_type="XML", word_size=6, expect=10, matrix_name="BLOSUM62", gapcosts="11 1")

with open('results.xml', 'w') as save_file: 
    blast_results = result_handle.read() 
    save_file.write(blast_results)

tree = ET.parse("results.xml")
root = tree.getroot()
result = []
for it in root.find("BlastOutput_iterations").findall("Iteration"):
    hits = it.find("Iteration_hits").findall("Hit")
    if len(hits) != 0:
        i = 0
        for hit in hits:
            hit_data = hit.find("Hit_hsps").find("Hsp")
            result.append(
            [it.find("Iteration_query-def").text, hit.find("Hit_id").text, hit_data.find("Hsp_bit-score").text, hit_data.find("Hsp_evalue").text]
            )
            i += 1
            if i == 3:
                break
    else:
        result.append([it.find("Iteration_query-def").text, "-", "-", "-"])


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('final_result.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Input ID")
worksheet.write(0, 1, "Nematostella ID")
worksheet.write(0, 2, "Score")
worksheet.write(0, 3, "eValue")


# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for row_data in (result):
    for col_data in row_data:
        worksheet.write(row, col, col_data)
        col += 1
    col = 0    
    row += 1

workbook.close()
