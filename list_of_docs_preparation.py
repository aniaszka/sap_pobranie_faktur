import pandas as pd
import logging.config
import logging

def prepare_doc_lists(excel_path, excel_name,
                      list_of_8010_docs_file_name,
                      list_of_8030_docs_file_name,
                      list_of_both_docs_file_name ):
    """
    funkcja wygeneruje 3 listy dokumentów - po jednej liście na jednostkę i listę łączną
    listy będą zapisane w plikach
    dodatkowo zwróci flagi z informacją, czy pliki nie są puste
    """
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file,
                              disable_existing_loggers=False)
    logger = logging.getLogger(__name__)
    logger.debug('Moduł do przygotowania plików txt uruchomiony.')

    #TODO ścieżka i nazwa pliku powinna być tylko raz zdefiniowana. teaz dubluej się z modułem enter_docs_sap_export

    file_full_name = excel_path + excel_name
    #TODO gdy export będzie na innym użytkowniku nazwy kolumn mogą być inne np pa angielsku
    docs_file = pd.read_excel(file_full_name, usecols=['Jednostka gosp.', 'Numer dokumentu'])

    df_8010 = docs_file[docs_file['Jednostka gosp.'] == 8010]
    df_8030 = docs_file[docs_file['Jednostka gosp.'] == 8030]
    df_both = docs_file['Numer dokumentu']

    if not df_8010.empty:
        series_8010 = df_8010['Numer dokumentu']
        series_8010.to_csv(excel_path + list_of_8010_docs_file_name, index=False, header=False)
        not_empty_8010 = True
    else:
        not_empty_8010 = False

    if not df_8030.empty:
        series_8030 = df_8030['Numer dokumentu']
        series_8030.to_csv(excel_path + list_of_8030_docs_file_name, index=False, header=False)
        not_empty_8030 = True
    else:
        not_empty_8030 = False

    if not df_both.empty:
        df_both.to_csv(excel_path + list_of_both_docs_file_name, index=False, header=False)
        not_empty_both = True
    else:
        not_empty_both = False

    return not_empty_8010, not_empty_8030, not_empty_both
    logger.debug(f'not_empty_8010: {not_empty_8010}, not_empty_8030: {not_empty_8030}, not_empty_both: {not_empty_both}')
