import requests


def clique_no_botao(bot, not_found):
    if not bot.find("gerar_ceps", matching=0.97, waiting_time=10000):
        not_found("gerar_ceps")
    bot.click()


def copiar_cep(botju, not_found, ceps_list):
    if not botju.find("copiarcep", matching=0.97, waiting_time=10000):
        not_found("copiarcep")
    botju.wait(500)
    botju.click()
    botju.control_a()
    botju.control_c()
    guarda = botju.get_clipboard()

    url = f'https://viacep.com.br/ws/{guarda}/json/'
    response = requests.get(url.format(guarda))
    response.raise_for_status()
    data = response.json()
    botju.wait(200)

    if 'erro' in data:
        print("O cep:", guarda, "não tem dados registrados na API")
    else:
        ceps_list.append(guarda)






