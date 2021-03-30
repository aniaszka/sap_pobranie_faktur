import logging.config
import logging


def generate_attachment_list(excel_path, list_of_both_docs_file_name, attachment_list_file, session):
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file, disable_existing_loggers=False)
    logger = logging.getLogger(__name__)
    logger.debug('Moduł do pobrania pliku z listą załączników uruchomiony.')

    session.StartTransaction(Transaction="Z052B_DOC_LIST")  # wywołanie transakcji
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[17]').Press() # otwarcie listy wariantów
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell').selectedRows = "0" # wskazanie wariantu
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[2]').Press() # potwierdzenie wyboru
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/btn%_S_BELNR_%_APP_%-VALU_PUSH').Press() # otwarcie okna z wyborem dokumentów
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[23]').Press() # klikamy opcję importu listy dok z pliku
    session.FindById('/app/con[0]/ses[0]/wnd[2]/usr/ctxtDY_PATH').text = excel_path
    session.FindById('/app/con[0]/ses[0]/wnd[2]/usr/ctxtDY_FILENAME').text = list_of_both_docs_file_name
    session.FindById('/app/con[0]/ses[0]/wnd[2]/usr/ctxtDY_FILENAME').caretPosition = 15
    session.FindById('/app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[8]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]').Press()
    # wybranie exportu danych
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/cntlCONTAINER1/shellcont/shell').pressToolbarContextButton("&MB_EXPORT")
    # wybranie formatu
    session.FindById('/app/con[0]/ses[0]/wnd[0]/usr/cntlCONTAINER1/shellcont/shell').selectContextMenuItem("&XXL")
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH').text = excel_path
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME').text = attachment_list_file
    session.FindById('/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME').caretPosition = 19
    session.FindById('/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[11]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]').Press()
    session.FindById('/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]').Press()

    logger.debug('Pobranie pliku z listą załączników wykonane.')

