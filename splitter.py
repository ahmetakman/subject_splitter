#Ahmet Akman 12/2019 splitter code

print (" started ")

from openpyxl import load_workbook


text_quantity = 5



column = ["A","B","C","D","E","F","G","H","I","J","K"]

checkpoints= ['Header1','Header2','Header3','Header4','Header5','Header6','Header7','Header8','Header9','sign']

"""
for i in range(text_quantity-1):
    t = open("path/text{}.txt".format(i+1),"a")
    t.write(" sign")
    t.close()
#/home/akman/Belgeler/diseases/newfolder/metinim1.txt
"""

for j in range(text_quantity-1):

    book = load_workbook(filename="final_folder.xlsx")

    sheet = book.active


    f = open("newfolder/metinim{}.txt".format(j+1),"r")
    
    whole_text = f.read()


    text = whole_text

    paster = [" ","","","","","","","","","","","","","",""]

    writer = text.split('header0')

    paster = [writer[0],*paster]

    last_situ = 1

    for q in range(10):
        writer1 = writer[1].split(checkpoints[q],1)
        if writer1[0] == writer[1]:
            writer1 = [" ",*writer1]

        else:
            paster[last_situ] = writer1[0]
            last_situ = q+2

        writer = writer1

    for p in range(11):
        sheet["{0}{1}".format(column[p],(j+1))] = paster[p]
    book.save(filename="final_folder.xlsx")
    book.close()


print("final : )")