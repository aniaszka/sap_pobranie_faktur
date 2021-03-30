import logging.config
import logging



def pobierz(jednostka, excel_path, session):
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file, disable_existing_loggers=False)
    logger = logging.getLogger(__name__)

    logger.debug('Uruchamiam transakcję do pobrania faktur.')

    session.StartTransaction(Transaction="ZFI_APA_REP")  # wywołanie transakcji

    #TODO może się zdarzyć, że przy otwartych oknach sapa zamiast ses[0] będzie inny numer
    #TODO wtedy poniższe makro nie zadziała.


    session.findById("wnd[0]/tbar[1]/btn[17]").Press()
    session.findById("wnd[1]/usr/txtV-LOW").text = "docs_year"
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[8]").Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/ctxtS_BUKRS-LOW').text = jednostka
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/txtS_BELNR-LOW').setFocus
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/txtS_BELNR-LOW').caretPosition = 0
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/btn%_S_BELNR_%_APP_%-VALU_PUSH').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[23]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[2]/usr/ctxtDY_PATH').setFocus
    session.FindById('/app/con[0]/ses[0]/wnd[2]/usr/ctxtDY_PATH').caretPosition = 0
    session.FindById('/app/con[0]/ses[0]/wnd[2]').sendVKey(4)

    #TODO zrobić pętlę bo teraz wykonuje skrypt tylko dla jednego pliku.


    #TODO ta ścieżka powtarza się wielokrotnie - poprawić np poprzez plik z konfiguracją

    session.FindById('/app/con[0]/ses[0]/wnd[3]/usr/ctxtDY_PATH').text = excel_path

    plik_z_lista = 'series_'+jednostka+'.txt'

    session.FindById('/app/con[0]/ses[0]/wnd[3]/usr/ctxtDY_FILENAME').text = plik_z_lista
    session.FindById('/app/con[0]/ses[0]/wnd[3]/usr/ctxtDY_FILENAME').caretPosition = 15
    session.FindById('/app/con[0]/ses[0]/wnd[3]/tbar[0]/btn[0]').Press() #todo czy ten wierszs jest potrzebny?
    session.FindById('/app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[8]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]').Press()

    # rozdzielam na dwie ścieżki, bo gdy nie ma dokumentów do pobrania to byłby błąd
    # flaga sukces jest potrzebna, żeby było wiadomoo, czy uruchomić osobny wątek na wskazanie miejsca pobrania faktur
    try:
        logger.debug('Klikam pobierz i będę czekać na wskazanie folderu.')
        session.FindById('/app/con[0]/ses[0]/wnd[0]').sendVKey(23)
        session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]').Press()
        session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]').Press()
        session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]').Press()
        logger.debug('Pobranie faktur zakończone sukcesem.')
        sukces = 1

    except:
        logger.debug('Z podanej listy dokumentów nie da się niczego pobrać.')
        session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").Press()
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]").Press()
        logger.debug('Wychodzę z transakcji.')
        sukces = 0

    return sukces



