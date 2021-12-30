"""
Matem
Busca los estudiantes admitidos en la UCR que llevaron Matem en los grupos de mate aplicada
Este programa SOLO se puede correr en SPYDER porque la librería xlrd opera con archivos .xls, los cuales excel ya no usa, pero por alguna razón en spyder sí funciona aún cuando en la consola NO.
Para que funcione siempre se puede usar pandas, pero me dio pereza hacerlo de nuevo jeje
"""

import xlrd

path = "C:\\Users\\shern\\Downloads\\"
path_matem = "C:\\Users\\shern\\Documents\\UCR\\Matem\\ADMITIDOS_MATEM_UCR_2021.xlsx"
ext = ".xlsx"

# inicializa el archivo de estudiantes admitidos
matem = xlrd.open_workbook(path_matem)
hoja_matem = matem.sheet_by_index(3) # 0: aprobados cálculo, 1: reprobados cálculo, 2: aprobados precálculo, 3: reprobados precálculo
carnets_adm = []
#print(hoja_matem.row(3))

# añade a una lista los carnés admitidos
for fila in range(4, hoja_matem.nrows):                 # al cambiar de hoja_matem es necesario actualizar la fila en que comienza
    carnets_adm.append(hoja_matem.cell_value(fila, 7))  # al cambiar de hoja_matem es necesario actualizar la columna
#print("Worksheet name(s): {0}".format(matem.sheet_names()))

correr = True

while correr:
    nombre = input("Nombre del archivo del curso: ")
    path_curso = path + nombre + ext
    print(path_curso)
    
    # abre los archivos excel y selecciona las hojas
    curso = xlrd.open_workbook(path_curso)
    hoja_curso = curso.sheet_by_index(0) # en el archivo del curso solo se usa la primera hoja
    
    comparar = 0
    estudiantes = 0
    aprobados = 0
    reprobados = 0
    total = 0
    
    # compara los carnés admitidos con los carnés del grupo (esto no cambia al cambiar de curso)
    for fila in range(1, hoja_curso.nrows):
        comparar = hoja_curso.cell_value(fila, 2)
        total += 1
        for carnet in carnets_adm:
            if carnet == comparar:
                estudiantes += 1
                print(hoja_curso.cell_value(fila, 1))
                if hoja_curso.cell_value(fila, 3) >= 7 or hoja_curso.cell_value(fila, 3) == "AP":
                    aprobados +=1
                elif hoja_curso.cell_value(fila, 3) < 7 or hoja_curso.cell_value(fila, 3) == "NAP":
                    reprobados += 1
    
    print("Estudiantes que llevaron Matem y el curso: ", estudiantes)
    print("Estudiantes aprobados: ", aprobados)
    print("Estudiantes reprobados: ", reprobados)
    print("Total de estudiantes en el curso: ", total)
    
    sigue = input("¿Comparar otro grupo? (0 = No)")
    if sigue == "0":
        correr = False

"""
for rx in range(hoja_matem.nrows):
    print(hoja_matem.row(rx))


print("The number of worksheets is {0}".format(curso.nsheets))
print("Worksheet name(s): {0}".format(curso.sheet_names()))
sh = curso.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
for rx in range(sh.nrows):
    print(sh.row(rx))
"""



    