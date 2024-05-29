
def atualizar_base(base, consulta):
    base_atualizada = base.copy()
    
    for index, row in consulta.iterrows():
        solicitacao = row['Solicitação']
        if solicitacao in base['Solicitação'].values:
            for col in consulta.columns:
                if col in base.columns:
                    if base.loc[base['Solicitação'] == solicitacao, col].values[0] != row[col]:
                        base_atualizada.loc[base_atualizada['Solicitação'] == solicitacao, col] = row[col]
        else:
            base_atualizada = pd.concat([base_atualizada, row.to_frame().T], ignore_index=True)
    
    return base_atualizada

