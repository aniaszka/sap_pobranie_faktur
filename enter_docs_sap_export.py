import logging.config
import logging
from datetime import datetime, timedelta

def pobierz_enter_docs(excel_path, excel_name, session):
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file, disable_existing_loggers=False)
    logger = logging.getLogger(__name__)
    logger.debug('Uruchamiam makro do pobrania dokumentów.')
    logger.debug('Dodatkowo funkcja ta da sygnał, czy są nowe dokumenty i czy dalsze kroki są potrzebne')

    session.StartTransaction(Transaction="FBL1N")  # wywołanie transakcji
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[17]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/txtENAME-LOW').text = "ASZADKOWSK"
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/txtENAME-LOW').setFocus
    session.FindById(
        '/app/con[0]/ses[0]/wnd[1]/usr/txtV-LOW').text = "OCR_IVOICES"
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/txtV-LOW').caretPosition = 11
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[8]').Press()
    # wiariant wybrany teraz czas na ustawienie daty
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[16]').Press()  # rozszerzam pola do filtrowania
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN014_%_APP_%-VALU_PUSH').Press()

    teraz = datetime.now()
    dzien_tygodnia = teraz.weekday()

    wczoraj = teraz - timedelta(days=1)
    wczoraj = wczoraj.strftime("%d.%m.%Y")

    przedwczoraj = teraz - timedelta(days=2)
    przedwczoraj = przedwczoraj.strftime("%d.%m.%Y")

    przedprzedwczoraj = teraz - timedelta(days=3)
    przedprzedwczoraj = przedprzedwczoraj.strftime("%d.%m.%Y")

    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]').text = wczoraj

    if dzien_tygodnia == 0:
        logger.debug('Do pobrania dokumenty z piątku, soboty, niedzeli.')
        session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]').text = przedwczoraj
        session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]').text = przedprzedwczoraj

    else:
        logger.debug('Generuje listę dokumentów z wczoraj.')

    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[8]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]').Press()

    # jeżeli w dniu wczorajszym nie było ksęgowań to dalsze kroki się wysypią dlatego wstawiam wyjątek
    try:
        session.FindById('/app/con[0]/ses[0]/wnd[0]/mbar/menu[0]/menu[3]/menu[1]').Select()
        session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]').Press()
        session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH').text = excel_path
        session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME').text = excel_name
        session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME').caretPosition = 10
        session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[11]").Press()
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]").Press()
        logger.debug('Lista dokumentów zapisana do pliku.')
        return True
    except:
        print('Brak danych z wczoraj. nie ma eksportu')
        logger.debug('Nie ma nowych dokumentów. Export się nie powiódł.')
        return False

    session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]").Press()


