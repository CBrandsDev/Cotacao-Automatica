import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pyautogui as py
import time

tabela = pd.read_excel('Cotação Base.xlsx')
tabela = tabela.drop("Dolar", axis=1)
print(tabela)

pergunta = input("Deseja atualizar a tabela de cotações? [Y/N]")
if pergunta == "n":
    quit()
else:

    navegador = webdriver.Chrome

    # segundo plano
    chrome_options = Options()
    chrome_options.headless = True 
    navegador = webdriver.Chrome(options=chrome_options)
    # Pegar cotação das moedas

    # peso argentino
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+peso+argentino&sxsrf=AJOqlzVikaNCBQO2v6CbVJ8SxuZMZ910_A%3A1673638507996&ei=a7LBY-y1POPQ1sQP1aeJiAg&oq=cota%C3%A7%C3%A3o+peso&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAxgAMg0IABCABBCxAxBGEIICMggIABCABBCxAzIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgcIIxDqAhAnOgcILhDqAhAnOgwIABDqAhC0AhBDGAE6BAgjECc6BAguECc6BAgAEEM6CggAELEDEIMBEEM6CwgAEIAEELEDEIMBOggIABCxAxCDAToHCAAQsQMQQ0oECEEYAEoECEYYAVAAWP31I2Cd_yNoA3AAeACAAZABiAG6DJIBBDAuMTKYAQCgAQGwARTAAQHaAQYIARABGAE&sclient=gws-wiz-serp')
    ct_pesoargentino = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # dolar australiano
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar+australiano&sxsrf=AJOqlzUOmWm2ALXSwYAOaWq_mRSIzj4_8g%3A1673639160736&ei=-LTBY4W_LJPM1sQP_OO00As&oq=cota%C3%A7%C3%A3o+dolar+aus&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAxgAMhAIABCABBCxAxCDARBGEIICMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOgcIABCwAxBDOhIILhDHARDRAxDIAxCwAxBDGAE6BAgjECc6CwgAEIAEELEDEIMBOggIABCABBCxAzoNCAAQgAQQsQMQgwEQCjoHCAAQgAQQCjoJCCMQJxBGEIICSgQIQRgASgQIRhgAUMcEWIQkYOwxaAJwAXgAgAGMAYgBqQqSAQQwLjEwmAEAoAEByAELwAEB2gEECAEYCA&sclient=gws-wiz-serp')
    ct_dlaustraliano = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')    
    # dolar canadense
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar+canadense&sxsrf=AJOqlzXPoSXf3_AqW8HvQTyPUq7O9kr3oA%3A1673639320569&ei=mLXBY6mtIsvU1sQP_56H2A8&oq=cota%C3%A7%C3%A3o+dolar+can&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAxgAMhAIABCABBCxAxCDARBGEIICMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOgcIABCwAxBDOgQIIxAnOggIABCABBDLAToJCCMQJxBGEIICOgsIABCABBCxAxCDAToICAAQgAQQsQNKBAhBGABKBAhGGABQhjxYzVpg7WJoAXABeACAAfQBiAHpDpIBBjAuMTMuMZgBAKABAcgBCsABAQ&sclient=gws-wiz-serp')
    ct_dlcanadense = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # franco suiço
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+franco+sui%C3%A7o&sxsrf=AJOqlzXbahSfUfM6nak9AXD-q_-dvkhwKQ%3A1673639561183&ei=ibbBY5zuCty75OUPyquGsAs&oq=cota%C3%A7%C3%A3o+fra&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAxgAMhAIABCABBCxAxCDARBGEIICMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOgcIABCwAxBDOhIILhDHARDRAxDIAxCwAxBDGAE6DAguEMgDELADEEMYAToLCAAQgAQQsQMQgwFKBAhBGABKBAhGGABQ2QNYzwZg5w5oAXABeACAAYIBiAH3ApIBAzAuM5gBAKABAcgBDcABAdoBBAgBGAg&sclient=gws-wiz-serp')
    ct_frsuico = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # Dolar
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar&sxsrf=AJOqlzXf5KCOQ8V8a8x_aE-khtct4M-hPw%3A1673639764859&ei=VLfBY-GLNKHb1sQPsI642AI&oq=cota%C3%A7%C3%A3o+dolar&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQARgCMgkIIxAnEEYQggIyBAgjECcyBAgjECcyCwgAEIAEELEDEIMBMggIABCABBCxAzILCAAQgAQQsQMQgwEyCAgAEIAEELEDMgUIABCABDILCAAQgAQQsQMQgwEyCwgAEIAEELEDEIMBOgoIABBHENYEELADOgcIABCwAxBDOgoIABCxAxCDARBDOgQIABBDSgQIQRgASgQIRhgAUL8FWJYIYI4UaAFwAXgAgAGIAYgB-gSSAQMwLjWYAQCgAQHIAQrAAQE&sclient=gws-wiz-serp')
    ct_dl = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # euro
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+euro&sxsrf=AJOqlzWIf6sRTmF8jRwmTFybH7X9W9dZZw%3A1673640022960&ei=VrjBY8aYOubx1sQPrYqqyAE&ved=0ahUKEwjGm7P4qsX8AhXmuJUCHS2FChkQ4dUDCBA&uact=5&oq=cota%C3%A7%C3%A3o+euro&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAzIJCCMQJxBGEIICMgsIABCABBCxAxCDATILCAAQgAQQsQMQgwEyBQgAEIAEMggIABCABBCxAzIICAAQgAQQsQMyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOg0IABBHENYEEMkDELADOggIABCSAxCwAzoSCC4QxwEQ0QMQyAMQsAMQQxgBOgQIIxAnSgQIQRgASgQIRhgAUNQFWOoHYPsSaAFwAHgAgAF9iAHwA5IBAzAuNJgBAKABAcgBC8ABAdoBBAgBGAg&sclient=gws-wiz-serp')
    ct_euro = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # libra esterlina
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+libra&sxsrf=AJOqlzUCPvydjsuO2v9Sdfy7xu0VObG0_w%3A1673640131451&ei=w7jBY_-GG6zI1sQPhqqFoA0&oq=cota%C3%A7%C3%A3o+libra&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQARgAMhAIABCABBCxAxCDARBGEIICMgsIABCABBCxAxCDATIFCAAQgAQyCwgAEIAEELEDEIMBMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOgcIABCwAxBDOhIILhDHARDRAxDIAxCwAxBDGAE6CAgAELEDEIMBSgQIQRgASgQIRhgAUNgEWL8JYIEUaAFwAXgAgAGFAYgB9gSSAQMwLjWYAQCgAQHIAQzAAQHaAQQIARgI&sclient=gws-wiz-serp')
    ct_lbet = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    # iene
    navegador.get(r'https://www.google.com/search?q=cota%C3%A7%C3%A3o+iene&sxsrf=AJOqlzVs6Q5IjotmCPybrmxm_HQE_WzKwg%3A1673640167572&ei=57jBY_7GIvb41sQP8tmMmAQ&oq=cota%C3%A7%C3%A3o+iene&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQARgAMhAIABCABBCxAxCDARBGEIICMggIABCABBCxAzIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEOgoIABBHENYEELADOgcIABCwAxBDOgQIIxAnOgsIABCABBCxAxCDAToJCCMQJxBGEIICOggIABCxAxCDAUoECEEYAEoECEYYAFDJBFiwE2CrHWgBcAF4AIABhgGIAfgIkgEDMC45mAEAoAEByAEKwAEB&sclient=gws-wiz-serp')
    ct_iene = navegador.find_element(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

#cotaçoes de criptos
# btc
    navegador.get(r'https://br.financas.yahoo.com/quote/BTC-USD/')
    ct_btc = navegador.find_element(
    'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
# eth
    navegador.get(r'https://finance.yahoo.com/quote/ETH-USD/')
    ct_eth = navegador.find_element(
    'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
# wbtc
    navegador.get(r'https://finance.yahoo.com/quote/WBTC-USD/')
    ct_wbtc = navegador.find_element(
    'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
# sol
    navegador.get(r'https://finance.yahoo.com/quote/SOL-USD/')
    ct_sol = navegador.find_element(
    'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
# ada
    navegador.get(r'https://finance.yahoo.com/quote/ADA-USD/')
    ct_ada = navegador.find_element(
        'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
    # xrp
    navegador.get(r'https://finance.yahoo.com/quote/XRP-USD/')
    ct_xrp = navegador.find_element(
        'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
    # bnb
    navegador.get(r'https://finance.yahoo.com/quote/BNB-USD/')
    ct_bnb = navegador.find_element(
        'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    
    # avax
    navegador.get(r'https://finance.yahoo.com/quote/AVAX-USD/')
    ct_avax = navegador.find_element(
        'xpath', '//*[@id="quote-header-info"]/div[3]/div[1]/div/fin-streamer[1]' ).get_attribute('value')
    

    # valor dos bens

    #ouro
    navegador.get(r'https://www.melhorcambio.com/ouro-hoje')
    ct_ouro = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_ouro = ct_ouro.replace(",", ".")
    
    # cafe
    navegador.get(r'https://www.melhorcambio.com/cafe-hoje')
    ct_cafe = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_cafe = ct_cafe.replace(",", ".")
    
    # petroleo
    navegador.get(r'https://www.melhorcambio.com/petroleo-hoje')
    ct_pt = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_pt = ct_pt.replace(",", ".")
    
    # soja
    navegador.get(r'https://www.melhorcambio.com/soja-hoje')
    ct_soja = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_soja = ct_soja.replace(",", ".")
    
    # etanol
    navegador.get(r'https://www.melhorcambio.com/etanol-hoje')
    ct_et = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_et = ct_et.replace(",", ".")
    
    # trigo
    navegador.get(r'https://www.melhorcambio.com/trigo-hoje')
    ct_trigo = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_trigo = ct_trigo.replace(",", ".")
    
    # boi 
    navegador.get(r'https://www.melhorcambio.com/boi-hoje')
    ct_boi = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_boi = ct_boi.replace(",", '.')
    
    # suino
    navegador.get(r'https://www.melhorcambio.com/suino-hoje')
    ct_suino = navegador.find_element(
        "xpath", '//*[@id="comercial"]').get_attribute('value')
    ct_suino = ct_suino.replace(",", ".")
    
    navegador.quit()

    #atualizar a tabela
    # moedas
    tabela.loc[tabela['Moedas'] == 'Peso Argentino', 'Moedas em Real'] = float(ct_pesoargentino)
    tabela.loc[tabela['Moedas'] == 'Peso Argentino', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Dólar Australiano', 'Moedas em Real'] = float(ct_dlaustraliano)
    tabela.loc[tabela['Moedas'] == 'Dólar Australiano', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Dólar Canadense', 'Moedas em Real'] = float(ct_dlcanadense)
    tabela.loc[tabela['Moedas'] == 'Dólar Canadense', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Franco Suíço', 'Moedas em Real'] = float(ct_frsuico)
    tabela.loc[tabela['Moedas'] == 'Franco Suíço', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Dólar Comercial', 'Moedas em Real'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Dólar Comercial', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Euro', 'Moedas em Real'] = float(ct_euro)
    tabela.loc[tabela['Moedas'] == 'Euro', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Libra Esterlina', 'Moedas em Real'] = float(ct_lbet)
    tabela.loc[tabela['Moedas'] == 'Libra Esterlina', 'Dolar'] = float(ct_dl)
    tabela.loc[tabela['Moedas'] == 'Iene', 'Moedas em Real'] = float(ct_iene)
    tabela.loc[tabela['Moedas'] == 'Iene', 'Dolar'] = float(ct_dl)
    # criptos
    tabela.loc[tabela['Criptos'] == 'BTC', 'Criptos em Dolar'] = float(ct_btc)
    tabela.loc[tabela['Criptos'] == 'ETH', 'Criptos em Dolar'] = float(ct_eth)
    tabela.loc[tabela['Criptos'] == 'WBTC', 'Criptos em Dolar'] = float(ct_wbtc)
    tabela.loc[tabela['Criptos'] == 'SOL', 'Criptos em Dolar'] = float(ct_sol)
    tabela.loc[tabela['Criptos'] == 'ADA', 'Criptos em Dolar'] = float(ct_ada)
    tabela.loc[tabela['Criptos'] == 'XRP', 'Criptos em Dolar'] = float(ct_xrp)
    tabela.loc[tabela['Criptos'] == 'BNB', 'Criptos em Dolar'] = float(ct_bnb)
    tabela.loc[tabela['Criptos'] == 'AVAX', 'Criptos em Dolar'] = float(ct_avax)
    # bens
    tabela.loc[tabela['Bens de consumo'] == 'Ouro', 'Valor em real'] = float(ct_ouro)
    tabela.loc[tabela['Bens de consumo'] == 'Cafe', 'Valor em real'] = str(ct_cafe)
    tabela.loc[tabela['Bens de consumo'] == 'Petróleo', 'Valor em real'] = float(ct_pt)
    tabela.loc[tabela['Bens de consumo'] == 'Soja', 'Valor em real'] = float(ct_soja)
    tabela.loc[tabela['Bens de consumo'] == 'Etanol', 'Valor em real'] = float(ct_et)
    tabela.loc[tabela['Bens de consumo'] == 'Trigo', 'Valor em real'] = str(ct_trigo)
    tabela.loc[tabela['Bens de consumo'] == 'Boi', 'Valor em real'] = float(ct_boi)
    tabela.loc[tabela['Bens de consumo'] == 'Suino', 'Valor em real'] = float(ct_suino)

    # print das cotações
    print('Cotação Peso Argentino:', ct_pesoargentino)
    print('Cotação Dolar Australiano:', ct_dlaustraliano)
    print('Cotação Dolar Canadense:', ct_dlcanadense)
    print('Cotação Franco Suiço:', ct_frsuico)
    print('Cotação Dolar:', ct_dl)
    print('Cotação Euro:', ct_euro)
    print('Cotação Libra Esterlina:', ct_lbet)
    print('Cotação Iene:', ct_iene)
    print('Cotação Bitcoin:', ct_btc)
    print('Cotação Etherium:', ct_eth)
    print('Cotação Wrapped Bitcoin:', ct_wbtc)
    print('Cotação Solana:', ct_sol)
    print('Cotação Cardano:', ct_ada)
    print('Cotação Ripple:', ct_xrp)
    print('Cotação Binance Coin:', ct_bnb)
    print('Cotação Avalanche:', ct_avax)
    print('Cotação Ouro:', ct_ouro)
    print('Cotação Cafe:', ct_cafe)
    print('Cotação Petróleo:', ct_pt)
    print('Cotação Soja:', ct_soja)
    print('Cotação Etanol:', ct_et)
    print('Cotação Trigo:', ct_trigo)
    print('Cotação Boi:', ct_boi)
    print('Cotação Suino:', ct_suino)

    # atualizaçao criptos em real
    tabela['Criptos em Real'] = tabela['Criptos em Dolar'] * tabela['Dolar']
    tabela = tabela.drop("Dolar", axis=1)
    print(tabela)
    time.sleep(2)
    print('Valores Atualizados')
    time.sleep(2)

    # exportar a tabela para novo arquivo
    #pergunta2 = input("Deseja exportar a nova tabela de cotações? [Y/N]")
    #if pergunta2== "n":
    #    quit()
    #else:
    tabela.to_excel('Cotação Atualizada.xlsx', index=False)
    print('Nova tabela exportada para a pasta destino como o nome Cotação Atualizada')
    time.sleep(3)
    print("Fechando progama em 3")
    time.sleep(1)
    print("2")
    time.sleep(1)
    print("1")
    time.sleep(1)
    #open(file= 'Cotação Atualizada.xlsx')
    quit()

    