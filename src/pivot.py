import win32com.client

def atualizar_tabelas(data_aging):

    # Abre o Modelo Aging em segundo plano
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    modelo_aging = excel.Workbooks.Open(rf"resultado\{data_aging}_Aging.xlsx")

    # Atualiza as Tabelas Dinâmicas e aguarda finalizar
    modelo_aging.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()

    analise = modelo_aging.Sheets('Aging (Etapas)')
    td_etapas = analise.PivotTables('TD_Etapas')

    td_etapas.PivotFields('ETAPA DA CONTA').ShowDetail = False

    # Remove campo da coluna 'EtapasTD'
    analise = modelo_aging.Sheets('Aging (Previsao)')
    td_analise = analise.PivotTables('TD_Analise')

    td_analise.PivotFields('EtapaTD').PivotItems("").Visible = False

    td_analise.PivotFields('EtapaTD').ShowDetail = False

    # Salva as alterações
    modelo_aging.save()
    modelo_aging.close()