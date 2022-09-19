import pandas as pd
from openpyxl import load_workbook

df = pd.read_csv('sp.csv')
dat = pd.DataFrame()

lt = {
    'date': [],
    'valor': [],
    'cod': [],
}

# input(':')
add_pl = load_workbook('matriz/matriz.xlsx')
a_at = add_pl.active

# input(': ')


def insert_vl(col, vlr):
    a_at[f'{col}'] = vlr


start_rel = 0
for pas in range(0, len(df)):
    if pas > 3:
        verifiq = df.loc[pas, 'Unnamed: 1']
        if str(verifiq).startswith('Banco'):
            start_rel += 1
            if start_rel == 1:
                d_saque = df.loc[pas, 'Declaração de conta - Shopee Cingapura']
                insert_vl('D3', f'{d_saque}')
            continue
        if start_rel == 1:
            date = df.loc[pas, 'Declaração de conta - Shopee Cingapura']
            lt['date'].append(date[:10])

            valor = df.loc[pas, 'Unnamed: 2']
            lt['valor'].append(valor)

            cod = df.loc[pas, 'Resumo']
            if str(cod).startswith('Renda do pedido'):
                lt['cod'].append(cod[18:])
            else:
                lt['cod'].append(cod)

tt_ped = len(lt['cod'])

st_mke = 11
tt_pd = 0

for pas in range(0, tt_ped):
    insert_vl(f'C{st_mke}', lt['cod'][pas])
    insert_vl(f'E{st_mke}', float(lt['valor'][pas]))
    insert_vl(f'B{st_mke}', str(lt['date'][pas]).replace('-', '/'))
    st_mke += 1
    tt_pd += 1

tt_pd -= 1

add_pl.save('Relat.xlsx')
print('-'*50)
print(f'Total de Pedidos: {tt_pd}')
print('Finalizado!')
print('-'*50)
print()
input('Pressione enter para sair!')