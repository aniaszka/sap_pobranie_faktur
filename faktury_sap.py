import sap_open
import enter_docs_sap_export
import list_of_docs_preparation as lodp
import attachment_list_generator
import invoices_download
import time
import win32com.client
import logging.config
import logging
import threading
import wskaz_folder
import sys
import lista_okien
import configparser


def faktury_sap():
    # logging_settings_file = Path()

    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file,
                              disable_existing_loggers=False)
    logger = logging.getLogger('faktury_sap')
    # nazwa bieżącego modułu (tylko w głównym poza tym __name__)

    logger.debug('Zaczynam program. Moduly zaimportowane')

    config = configparser.ConfigParser(interpolation=None)
    config.read('sap_pobranie_faktur_settings.ini')

    # todo sprawdzić, czy zmienne nie powtarzają się w innych modułach
    excel_name = "enter_docs.XLSX"

    excel_path = config['server_path']['excel_path'] # string

    logger.debug('Sciezki wskazane.')

    list_of_8010_docs_file_name = 'series_8010.txt'
    list_of_8030_docs_file_name = 'series_8030.txt'
    list_of_both_docs_file_name = 'series_both.txt'
    # plik użyty przy pobieraniu listy załączników do dokumentów

    attachment_list_file = "sap_attachments.XLSX"

    # 1. otwarcie sapa

    # sap_open.sap_open() # otwieram w osobnym wątku,
    # bo się lubi zawieszać i blokuje program
    my_thread = threading.Thread(target=sap_open.sap_open, daemon=True)
    my_thread.start()

    time.sleep(10)

    logger.debug('SAP powinien już być otwarty.')

    # połączenie z SAPem
    session = None
    logger.debug(f'Przed połączniem: session {session}.')
    while session == None:
        try:
            session = None
            SapGui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
            session = SapGui.FindById('ses[0]')  # połączenie z SAPem
            logger.debug(f'Mam połączenie: session {session}.')

        except:
            time.sleep(1)
            logger.debug(f'Ciągle nie ma połączenia {session}')

    # Połączenie z SAPem nawiązene więc zamykam chrome
    lista_okien.close_window_by_part_name('Bez tytułu - Google Chrome')

    # 2. pobranie pliku enter_docs.XLSX

    flaga = enter_docs_sap_export.pobierz_enter_docs(excel_path,
                                                     excel_name,
                                                     session)
    # flaga daje informację, czy są nowe dokumuenty.
    # Jeżeli nie to pozostaje jedynie zamknąć SAPa.

    if flaga:
        logger.debug('Są faktury, więc eksportuję je do pliku.')
        # 3. zamknięcie pliku excel
        time.sleep(10)
        logger.debug('Próbuję zamknąć plik excel enter_docs.')
        flaga = False
        i = 0
        while flaga == False:
            flaga = lista_okien.close_window_by_part_name('enter_docs')
            time.sleep(5)
            i += 1
            print(f'Próba numer: {i}')
            print(flaga)
            if flaga == True:
                print('Koniec pętli, okno excela odnalezione.')
                break
            if i == 10:
                print('Przerywam pętlę po 10 próbach.')
                break



        # 4. pliki txt z listami dokumentów
        logger.debug('Z danych w excelu zrobę listy dokumentów w plikach txt.')
        (not_empty_8010,
         not_empty_8030,
         not_empty_both) = lodp.prepare_doc_lists(excel_path,
                                                  excel_name,
                                                  list_of_8010_docs_file_name,
                                                  list_of_8030_docs_file_name,
                                                  list_of_both_docs_file_name)

        # 5. wygenerowanie pliku z listą załączników
        if not_empty_both:
            attachment_list_generator.generate_attachment_list(excel_path,
                                                list_of_both_docs_file_name,
                                                attachment_list_file,
                                                session)
        else:
            print('Nie ma dziś żadnych nowych dokumentów')

        # 6. zamknięcie pliku z listą załączników
        time.sleep(10)
        logger.debug('Próbuję zamknąć plik excel sap_attachments.')
        flaga = False
        i = 0
        while flaga == False:
            flaga = lista_okien.close_window_by_part_name('sap_attachments')
            time.sleep(5)
            i += 1
            print(f'Próba numer: {i}')
            print(flaga)
            if flaga == True:
                print('Koniec pętli, okno excela odnalezione')
                break
            if i == 10:
                print('Przerywam pętlę po 10 próbach.')
                break


        # 7. pobranie obrazów faktur

        if not_empty_8030:
            time.sleep(2)
            petla_folder = threading.Thread(target=wskaz_folder.wstaw_sciezke,
                                            args=())

            petla_folder.start()
            # pobranie.start()
            # join musi być poniżej wywołania transakcji
            # transakcja musi być wywołana bezpośrednio w głównym wątku,
            # a nie przez threading, bo SAP blokuje dostęp.
            # nie zawsze tak się dzieje, ale może.
            sukces = invoices_download.pobierz('8030', excel_path, session)
            print('skonczyłam z 8030. jestem tutaj.')
            if sukces == 1:
                petla_folder.join()


            # pobranie.join()
            # invoices_download.pobierz('8030', excel_path, session)
                time.sleep(2)


        else:
            print('Nie ma czego pobierać dla 8030')

        if not_empty_8010:
            print('przechodze do 8010')
            petla_folder = threading.Thread(target=wskaz_folder.wstaw_sciezke,
                                            args=())
            petla_folder.start()
            # pobranie.start()
            sukces = invoices_download.pobierz('8010', excel_path, session)
            if sukces == 1:
                petla_folder.join()
            # pobranie.join()
                time.sleep(2)

        else:
            print('Nie ma czego pobierać dla 8010')


        # 8 zamknięcie SAPa
        session.findById("/app/con[0]/ses[0]/wnd[0]").Close()
        session.findById("/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1").Press()
        time.sleep(5)
        lista_okien.close_window_by_part_name('SAP Logon 740')
        logger.debug('SAP zamknięty. Poszło wspaniale.\n')

    else:
        logger.debug('Nie ma nowych dokumentów. Kończę na dziś.')
        # 8 zamknięcie SAPa
        session.findById("/app/con[0]/ses[0]/wnd[0]").Close()
        session.findById("/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1").Press()
        time.sleep(5)
        lista_okien.close_window_by_part_name('SAP Logon 740')
        logger.debug('Zamykam SAPa. Koniec procesu.\n')

    print('Koniec procesu.')
    sys.exit
    # dodaję tę komendę, żeby wyłączyć
    # daemona - wątek otwarcia SAPa i jego zawieszania.

if __name__ == '__main__':
    faktury_sap()