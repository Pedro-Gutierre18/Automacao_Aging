from datetime import datetime, timedelta

def data_atual():
    data = datetime.now()

    data_ontem = data - timedelta(days=1)
    return data_ontem


def formatar_data(data_ontem):
    # Formata a data
    data_formatada = data_ontem.strftime("%d/%m/%Y")
    return data_formatada


def formatar_data_aging(data_ontem):
    # Formata a data para criação do nome padrão do arquivo gerado
    data_aging = data_ontem.strftime("%Y.%m.%d")
    return data_aging