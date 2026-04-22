import pandas as pd
import matplotlib.pyplot as plt
from decimal import *
def _letter_to_number(letter):
    letter = letter.upper().strip()
    column_number = 0
    for i, char in enumerate(reversed(letter)):
        char_value = (ord(char) - 64) * (26 ** i)
        column_number += char_value
    return column_number - 1
def excel_range_to_array(range_address_1):
    start_cell_1, end_cell_1 = range_address_1.split(':')
    start_column_1 = _letter_to_number(start_cell_1[0])
    start_row_1 = int(''.join(filter(str.isdigit, start_cell_1))) - 1
    end_column_1 = _letter_to_number(end_cell_1[0])
    end_row_1 = int(''.join(filter(str.isdigit, end_cell_1))) - 1
    selected_range_1 = excel_data.iloc[start_row_1:end_row_1+1, start_column_1:end_column_1+1]
    result_array_1 = selected_range_1.values.tolist()
    return result_array_1
ee0 = Decimal(input("Введите диэлектрическую проницаемость вашего материала _"))*Decimal('8.85418781762039e-12') 
S = Decimal(input("Введите площадь поперечного сечения _"))
qe = Decimal('-1.60217e-19')
const = 2/(ee0*qe*S**2)
file_path = input("Введите путь к Файлу с измерениями :_")
sheet_name = int(input("Введите номер листа с измерениями :_"))-1
excel_data = pd.read_excel(file_path, sheet_name, header = None, engine = 'openpyxl')
range_address_1 = input("Введите диапазон с измерениями C (ёмкости) _")
Capit_izm = excel_range_to_array(range_address_1)
range_address_2 = input("Введите диапазон с измерениями V (обратного смещения) _")
V_izm = excel_range_to_array(range_address_2)
Capit_obr_k = [' ']*len(Capit_izm)     
X_mass_zn = [' ']*len(Capit_izm)       
N_mass = [' ']*len(Capit_izm)      
for i in range(0, len(Capit_izm)):
    Capit_obr_k[i] = Decimal(str(Capit_izm[i][0])+'e-9')**(-2)
    X_mass_zn[i] = ee0*S/Decimal(str(Capit_izm[i][0])+'e-9')
for i in range(len(Capit_obr_k)-1):
    N_mass[i] = const/((Decimal(str(Capit_obr_k[i+1])) - Decimal(str(Capit_obr_k[i])))/(Decimal(str(V_izm[i+1][0])) - Decimal(str(V_izm[i][0]))))   
N_mass[-1] = N_mass[-2]
Dots_for_Graph = [' ']*len(Capit_izm)
for i in range(0, len(Capit_izm)):
    Dots_for_Graph[i] = [X_mass_zn[i], N_mass[i]]
X_nada_schitat = Decimal(input(f"Измерения какой глубины вам нужны? от {min(X_mass_zn)} до {max(X_mass_zn)} :_"))
N_nada_schitat = 'Вне измерений'
for i in range(len(X_mass_zn)-1):
    if X_nada_schitat < X_mass_zn[i] and X_nada_schitat >= X_mass_zn[i+1]:
        N_nada_schitat = N_mass[i]
print(f"Введенной глубине соответствует концентрация {N_nada_schitat}")
x = [nsns[0] for nsns in Dots_for_Graph]
y = [nsns[1] for nsns in Dots_for_Graph]
plt.figure(figsize=(15 , 35))
plt.plot(x, y, color='blue', linewidth=1, marker='o', markersize=2, markerfacecolor='red', markeredgecolor='darkred',markeredgewidth=2, label='Координаты')
plt.title('График зависимоти N от X', fontsize=16)
plt.xlabel('X, глубина измерений, м', fontsize=12)
plt.ylabel('N, концентрация свободных нчителей заряда, м^(-3)', fontsize=12)
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()
plt.show()
