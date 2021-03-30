import webbrowser
import logging.config
import logging
import configparser


def sap_open():
    # otwarcie sesji
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file,
                              disable_existing_loggers=False)
    logger = logging.getLogger(__name__)

    logger.debug('uruchamiam sap_open.')

    config = configparser.ConfigParser(interpolation=None)
    config.read('sap_pobranie_faktur_settings.ini') # nazwa pliku
    chrome_path = config.get('chrome_path', 'chrome_path')
    # nazwa sekcji, nazwa zimiennej
    logger.debug('Ścieżkę do chrome pobieram do pliku sap_pobranie_faktur_settings.ini')
    sap_link = config['sap']['sap_link']
    webbrowser.get(chrome_path).open_new_tab(sap_link)
    logger.debug('Strona SAPa odpalona.')


if __name__ == '__main__':
    sap_open()