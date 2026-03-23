import pandas as pd
from datetime import datetime
from integracoes import (
    get_access_token, get_users_and_groups, carregar_regras_rh, 
    carregar_aprovacoes_he, obter_cartoes_ponto, carregar_feriados,
    carregar_ferias, upload_excel_sharepoint
)
from motor_calculo import process_timecards

def main():
    print("--- INICIANDO MOTOR VALESHOP ---")
    
    # 1. Autenticação
    tk = get_access_token()
    if not tk:
        print("❌ Erro ao obter token do Microsoft Graph.")
        return

    # 2. Extração (Extract)
    mapa = get_users_and_groups(tk)
    regras_contratos = carregar_regras_rh(tk)
    aprovacoes_he = carregar_aprovacoes_he(tk)
    timecards = obter_cartoes_ponto(tk)
    feriados = carregar_feriados(tk)
    ferias = carregar_ferias(tk)

    # 3. Transformação (Transform)
    linhas = process_timecards(timecards, mapa, regras_contratos, aprovacoes_he, feriados, ferias)
    
    # A LINHA QUE FALTAVA: Transformar a lista do Python em uma tabela do Pandas
    df = pd.DataFrame(linhas)

    # 4. Carga (Load) - Salvando o arquivo Excel localmente
    nome_do_arquivo = "Relatorio_Ponto_ValeShop.xlsx"
    df.to_excel(nome_do_arquivo, index=False)
    print(f"Arquivo local '{nome_do_arquivo}' gerado com sucesso!")

    # 🚀 O GRAN FINALE: Enviando para a nuvem!
    # Nota: Crie a pasta "Relatorios_Ponto" dentro de Documentos no site do RH.
    upload_excel_sharepoint(tk, nome_do_arquivo, "Relatorios_Ponto")
    
    print("\n🎉 AUTOMAÇÃO FINALIZADA COM SUCESSO! 🎉")

if __name__ == "__main__":
    main()
