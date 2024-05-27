import pandas as pd

def calcular_status(row):
    data_sistema = pd.Timestamp.now()
    diferenca = (data_sistema - row['Data Retorno']).days
    if row['Data Retorno'] > data_sistema:
        return 'No Prazo'
    elif diferenca <= 5:
        return 'Px Prazo Final'
    else:
        return 'Em Atraso'
    
