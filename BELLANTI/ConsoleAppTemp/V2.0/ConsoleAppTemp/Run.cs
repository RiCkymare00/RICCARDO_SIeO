using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;

namespace ConsoleAppTemp
{
    internal class Run
    {
        public void ExecProgram(string tipo)
        {
            FTP ftp = new();
            if (string.IsNullOrWhiteSpace(tipo))
            {
                _ = new LogWriter("Nel file di config il tipo di app non risulta corretto. Selezionarne uno valido.");
            }
            try
            {
                //file da mandare a Bellanti
                if (tipo.Equals("1"))
                {
                    // Carica i dati dal database
                    string selQry = @"select * 
            from bellanti_match_articoli_gcc 
            where cod_bellanti is null 
              and len(cod_copre) > 0 
              and DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
              and ordinabile_copre = 's'
            ";
                    DataTable dt = new DataTable();
                    DBAccess.ReadDataThroughAdapter(selQry, dt);

                    // Carica i dati dal file Excel con le descrizioni dei livelli
                    var workbook = new ExcelFile();
                    workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                    var worksheet = workbook.Worksheets[0];
                    var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                        .ToDictionary(
                            row => row.Cells[0].Value.ToString(),
                            row => row.Cells[1].Value.ToString()
                        );

                    // Crea un nuovo file Excel per i risultati
                    var newWorkbook = new ExcelFile();
                    var newWorksheet = newWorkbook.Worksheets.Add("Results");

                    // Aggiungi le intestazioni
                    newWorksheet.Cells[0, 0].Value = "Codice Prodotto";
                    newWorksheet.Cells[0, 1].Value = "Descrizione";
                    newWorksheet.Cells[0, 2].Value = "Codice livello 1";
                    newWorksheet.Cells[0, 3].Value = "Descrizione Livello 1";
                    newWorksheet.Cells[0, 4].Value = "Codice livello 2";
                    newWorksheet.Cells[0, 5].Value = "Descrizione Livello 2";
                    newWorksheet.Cells[0, 6].Value = "Codice livello 3";
                    newWorksheet.Cells[0, 7].Value = "Descrizione Livello 3";
                    newWorksheet.Cells[0, 8].Value = "Codice livello 4";
                    newWorksheet.Cells[0, 9].Value = "Descrizione Livello 4";
                    newWorksheet.Cells[0, 10].Value = "Aliquota";
                    newWorksheet.Cells[0, 11].Value = "Modello";
                    newWorksheet.Cells[0, 12].Value = "Codice Marchio";
                    newWorksheet.Cells[0, 13].Value = "Descrizione Marchio";
                    newWorksheet.Cells[0, 14].Value = "EAN";
                    newWorksheet.Cells[0, 15].Value = "Prezzo Acquisto";
                    newWorksheet.Cells[0, 16].Value = "Prezzo Vendita";

                    int rowIndex = 1;

                    foreach (DataRow row in dt.Rows)
                    {
                        string codCopre = row["cod_copre"].ToString();
                        string catEdielCopre = row["cat_ediel_copre"].ToString();
                        string desArtCopre = row["des_art_copre"].ToString();
                        string modello = ConvertToString(row["modello"]);
                        string codiceMarchio = ConvertToString(row["cod_brand"]);
                        string descrizioneMarchio = ConvertToString(row["des_brand"]);
                        string aliquota = ConvertToString(row["cod_iva"]);
                        string ean = ConvertToString(row["ean"]);
                        string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                        string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);

                        // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                        string codiceL1, descrizioneL1;
                        string codiceL2, descrizioneL2;
                        string codiceL3, descrizioneL3;
                        string codiceL4, descrizioneL4;

                        // Livello 1
                        GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                        // Livello 2
                        if (catEdielCopre.Length >= 4)
                        {
                            GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                        }
                        else
                        {
                            codiceL2 = "XX";
                            descrizioneL2 = "Livello Generico";
                        }

                        // Livello 3
                        if (catEdielCopre.Length >= 6)
                        {
                            GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                        }
                        else
                        {
                            codiceL3 = "XX";
                            descrizioneL3 = "Livello Generico";
                        }

                        // Livello 4
                        if (catEdielCopre.Length == 8)
                        {
                            GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                        }
                        else
                        {
                            codiceL4 = "XX";
                            descrizioneL4 = "Livello Generico";
                        }

                        // Scrivi i dati nella nuova riga del file Excel
                        newWorksheet.Cells[rowIndex, 0].Value = codCopre;
                        newWorksheet.Cells[rowIndex, 1].Value = desArtCopre;
                        newWorksheet.Cells[rowIndex, 2].Value = codiceL1;
                        newWorksheet.Cells[rowIndex, 3].Value = descrizioneL1;
                        newWorksheet.Cells[rowIndex, 4].Value = codiceL2;
                        newWorksheet.Cells[rowIndex, 5].Value = descrizioneL2;
                        newWorksheet.Cells[rowIndex, 6].Value = codiceL3;
                        newWorksheet.Cells[rowIndex, 7].Value = descrizioneL3;
                        newWorksheet.Cells[rowIndex, 8].Value = codiceL4;
                        newWorksheet.Cells[rowIndex, 9].Value = descrizioneL4;
                        newWorksheet.Cells[rowIndex, 10].Value = aliquota;
                        newWorksheet.Cells[rowIndex, 11].Value = modello;
                        newWorksheet.Cells[rowIndex, 12].Value = codiceMarchio;
                        newWorksheet.Cells[rowIndex, 13].Value = descrizioneMarchio;
                        newWorksheet.Cells[rowIndex, 14].Value = ean;
                        newWorksheet.Cells[rowIndex, 15].Value = prezzoAcquisto;
                        newWorksheet.Cells[rowIndex, 16].Value = prezzoVendita;

                        rowIndex++;
                    }

                    // Salva il nuovo file Excel
                    newWorkbook.SaveXlsx(ConfigurationManager.AppSettings["OutputPathTipo1"]);
                }

                //creo il file Items e il file Stock come il tracciato di Miele. Stock ogni ora, Items una volta al giorno
                if (tipo.Equals("2") || tipo.Equals("3"))
                {

                    

                    //creo il file Items una volta al giorno
                    if (tipo.Equals("2"))
                    {
                        // Invio email di inizio processo
                        InviaEmail("MIA - Inizio aggiornamento articoli COPRE", "Il processo di aggiornamento articoli COPRE è iniziato.");

                        Console.WriteLine("Eseguo la query per prendere gli articoli  di Copre");
                        // Carica i dati dal database
                        string selQry = @"WITH CTE_Latest AS (
                            SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
                                   cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
                                   flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
                                   flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
                                   flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
                                   prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
                                   ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
                                   ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
                                   ultima_presenza_nel_file_copre,
                                   ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
                            FROM bellanti_match_articoli_gcc 
                            WHERE cod_bellanti IS NULL 
                              AND LEN(cod_copre) > 0 
                              AND DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
                              AND ordinabile_copre = 's'
                              AND LEN(ean) <= 13
                              AND ean IS NOT NULL
                              AND LTRIM(RTRIM(ean)) <> ''
                              AND prz_acquisto_copre > 0
                        ),
                        CTE_FirstBrand AS (
                            SELECT cod_brand, MIN(des_brand) AS des_brand
                            FROM CTE_Latest
                            GROUP BY cod_brand
                        ),
                        CTE_Filtered AS (
                            SELECT c.*
                            FROM CTE_Latest c
                            JOIN CTE_FirstBrand fb
                            ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
                        )
                        SELECT *
                        FROM CTE_Filtered
                        WHERE rn = 1;
                        ";

                        //                        string selQry = @"WITH CTE_Latest AS (
                        //    SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
                        //           cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
                        //           flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
                        //           flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
                        //           flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
                        //           prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
                        //           ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
                        //           ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
                        //           ultima_presenza_nel_file_copre,
                        //           ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
                        //    FROM bellanti_match_articoli_gcc 
                        //    WHERE cod_bellanti IS NULL 
                        //      AND LEN(cod_copre) > 0 
                        //      AND ultima_presenza_nel_file_copre IN ('20250314')
                        //      AND ordinabile_copre = 's'
                        //      AND LEN(ean) <= 13
                        //      AND ean IS NOT NULL
                        //      AND LTRIM(RTRIM(ean)) <> ''
                        //      AND prz_acquisto_copre > 0
                        //),
                        //CTE_FirstBrand AS (
                        //    SELECT cod_brand, MIN(des_brand) AS des_brand
                        //    FROM CTE_Latest
                        //    GROUP BY cod_brand
                        //),
                        //CTE_Filtered AS (
                        //    SELECT c.*
                        //    FROM CTE_Latest c
                        //    JOIN CTE_FirstBrand fb
                        //    ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
                        //)
                        //SELECT *
                        //FROM CTE_Filtered
                        //WHERE rn = 1;
                        //";

                        DataTable dt = new DataTable();
                        DBAccess.ReadDataThroughAdapter(selQry, dt);

                        Console.WriteLine("Carico il file excel dei livelli di Copre");
                        // Carica i dati dal file Excel con le descrizioni dei livelli
                        var workbook = new ExcelFile();
                        workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                        var worksheet = workbook.Worksheets[0];
                        var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                            .ToDictionary(
                                row => row.Cells[0].Value.ToString(),
                                row => row.Cells[1].Value.ToString()
                            );

                        // Crea un nuovo file CSV per i risultati
                        Console.WriteLine("Creo un nuovo file CSV per scrivere gli articoli");
                        string outputPath = ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
                        using (StreamWriter sw = new StreamWriter(outputPath))
                        {
                            // Aggiungi le intestazioni
                            string[] headers = new string[] {
                    "Codice_Prodotto", "Codice_EAN", "Prezzo_Acquisto", "Prezzo_Vendita", "Prezzo_Listino",
                    "Spese_Consegna", "Spese_Consegna_Incluse", "Attivo", "Codice_Fornitore", "Codice_Cliente",
                    "Descrizione_Breve", "Descrizione_Estesa", "Data_Lancio", "Classificazione", "Descrizione_Classificazione",
                    "Codice_Marchio", "Descrizione_Marchio", "Iva", "Note", "Link_Scheda_Energetica",
                    "Link_Video", "Data_Fuori_Produzione", "Raee", "Raee Annegata", "Siae", "Siae Annegata"
                };
                            sw.WriteLine(string.Join(";", headers));

                            Console.WriteLine("Inizio a ciclare gli articoli. Gli articoli sono: " + dt.Rows.Count);
                            int count = 0;
                            // Aggiungi i dati
                            foreach (DataRow row in dt.Rows)
                            {
                                count++;
                                Console.WriteLine("Articolo " + count + " di " + dt.Rows.Count);
                                string codCopre = row["cod_copre"].ToString();
                                string catEdielCopre = row["cat_ediel_copre"].ToString();
                                string desArtCopre = row["des_art_copre"].ToString();
                                string modello = ConvertToString(row["modello"]);
                                string codiceMarchio = ConvertToString(row["cod_brand"]);
                                string descrizioneMarchio = ConvertToString(row["des_brand"]);
                                string aliquota = "0";
                                string ean = ConvertToString(row["ean"]);
                                string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                                string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);

                                //// Gestione dell'aliquota IVA
                                //if (aliquota == "NOR")
                                //{
                                //    aliquota = "22";
                                //}
                                //else
                                //{
                                //    aliquota = Convert.ToInt32(Convert.ToDecimal(aliquota)).ToString();
                                //}

                                // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                                string codiceL1, descrizioneL1;
                                string codiceL2, descrizioneL2;
                                string codiceL3, descrizioneL3;
                                string codiceL4, descrizioneL4;

                                // Livello 1
                                GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                                // Livello 2
                                if (catEdielCopre.Length >= 4)
                                {
                                    GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                                }
                                else
                                {
                                    codiceL2 = "XX";
                                    descrizioneL2 = "Livello Generico";
                                }

                                // Livello 3
                                if (catEdielCopre.Length >= 6)
                                {
                                    GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                                }
                                else
                                {
                                    codiceL3 = "XX";
                                    descrizioneL3 = "Livello Generico";
                                }

                                // Livello 4
                                if (catEdielCopre.Length == 8)
                                {
                                    GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                                }
                                else
                                {
                                    codiceL4 = "XX";
                                    descrizioneL4 = "Livello Generico";
                                }

                                // Scrivi i dati nella nuova riga del file CSV
                                string[] data = new string[] {
                        codCopre, ean, prezzoAcquisto, prezzoVendita, prezzoVendita,
                        "0", "0", "1", "", "",
                        desArtCopre.Replace("\"","\"\""), "", "", codiceL2, descrizioneL2.Replace("\"","\"\""),
                        codiceMarchio, descrizioneMarchio.Replace("\"","\"\""), aliquota, "", "",
                        "", "", "0", "0", "0", "0"
                    };
                                sw.WriteLine(string.Join(";", data.Select(d => $"\"{d}\"")));
                            }
                        }

                        ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                        Console.WriteLine("File CSV generato con successo: " + outputPath);

                        // Invio email di fine processo
                        InviaEmail("MIA - Fine aggiornamento articoli COPRE", "Il processo di aggiornamento articoli COPRE è terminato con successo.");

                    }

                    //creo il file Stock ogni ora
                    if (tipo.Equals("3"))
                    {
                        InviaEmail("MIA - Inizio aggiornamento giacenze COPRE", "Il processo di aggiornamento giacenze COPRE è iniziato.");
                        Console.WriteLine("Eseguo la query per prendere le giacenze  di Copre");
                        // Carica i dati dal database
                        string selQry2 = @"WITH CTE_Latest AS (
    SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
           cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
           flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
           flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
           flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
           prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
           ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
           ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
           ultima_presenza_nel_file_copre,
           ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
    FROM bellanti_match_articoli_gcc 
    WHERE cod_bellanti IS NULL 
      AND LEN(cod_copre) > 0 
      AND DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
      AND ordinabile_copre = 's'
      AND LEN(ean) <= 13
      AND ean IS NOT NULL
      AND LTRIM(RTRIM(ean)) <> ''
      AND prz_acquisto_copre > 0
),
CTE_FirstBrand AS (
    SELECT cod_brand, MIN(des_brand) AS des_brand
    FROM CTE_Latest
    GROUP BY cod_brand
),
CTE_Filtered AS (
    SELECT c.*
    FROM CTE_Latest c
    JOIN CTE_FirstBrand fb
    ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
)
SELECT *
FROM CTE_Filtered
WHERE rn = 1;
";
                        DataTable dt2 = new DataTable();
                        DBAccess.ReadDataThroughAdapter(selQry2, dt2);

                        Console.WriteLine("Carico il file excel dei livelli di Copre");
                        // Carica i dati dal file Excel con le descrizioni dei livelli
                        var workbook = new ExcelFile();
                        workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                        var worksheet = workbook.Worksheets[0];
                        var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                            .ToDictionary(
                                row => row.Cells[0].Value.ToString(),
                                row => row.Cells[1].Value.ToString()
                            );

                        // Crea un nuovo file Excel per i risultati
                        Console.WriteLine("Creo un nuovo file CSV per scrivere le giacenze");
                        var newWorkbook = new ExcelFile();
                        var newWorksheet = newWorkbook.Worksheets.Add("Results");

                        // Aggiungi le intestazioni
                        newWorksheet.Cells[0, 0].Value = "Codice_Prodotto";
                        newWorksheet.Cells[0, 1].Value = "Codice_EAN";
                        newWorksheet.Cells[0, 2].Value = "Quantita_Disponibile";
                        newWorksheet.Cells[0, 3].Value = "Quantita_in_Arrivo";
                        newWorksheet.Cells[0, 4].Value = "Data_Arrivo";
                        newWorksheet.Cells[0, 5].Value = "Codice_Fornitore";
                        newWorksheet.Cells[0, 6].Value = "Codice_Cliente";

                        int rowIndex = 1;

                        int count = 0;

                        Console.WriteLine("Inizio a ciclare le giacenze. Le giacenze sono: " + dt2.Rows.Count);
                        foreach (DataRow row in dt2.Rows)
                        {
                            count++;
                            Console.WriteLine("Giacenza " + count + " di " + dt2.Rows.Count);
                            string codCopre = row["cod_copre"].ToString();
                            string catEdielCopre = row["cat_ediel_copre"].ToString();
                            string desArtCopre = row["des_art_copre"].ToString();
                            string modello = ConvertToString(row["modello"]);
                            string codiceMarchio = ConvertToString(row["cod_brand"]);
                            string descrizioneMarchio = ConvertToString(row["des_brand"]);
                            string aliquota = "22";
                            string ean = ConvertToString(row["ean"]);
                            string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                            string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);
                            string qtaInArrivo = ConvertToInt(row["ordine_fornitore_copre"]);
                            string qtaDisponibile = ConvertToInt(row["giacenza_copre"]);

                            // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                            string codiceL1, descrizioneL1;
                            string codiceL2, descrizioneL2;
                            string codiceL3, descrizioneL3;
                            string codiceL4, descrizioneL4;

                            // Livello 1
                            GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                            // Livello 2
                            if (catEdielCopre.Length >= 4)
                            {
                                GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                            }
                            else
                            {
                                codiceL2 = "XX";
                                descrizioneL2 = "Livello Generico";
                            }

                            // Livello 3
                            if (catEdielCopre.Length >= 6)
                            {
                                GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                            }
                            else
                            {
                                codiceL3 = "XX";
                                descrizioneL3 = "Livello Generico";
                            }

                            // Livello 4
                            if (catEdielCopre.Length == 8)
                            {
                                GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                            }
                            else
                            {
                                codiceL4 = "XX";
                                descrizioneL4 = "Livello Generico";
                            }

                            // Scrivi i dati nella nuova riga del file Excel
                            newWorksheet.Cells[rowIndex, 0].Value = codCopre;
                            newWorksheet.Cells[rowIndex, 1].Value = ean;
                            newWorksheet.Cells[rowIndex, 2].Value = qtaDisponibile;
                            newWorksheet.Cells[rowIndex, 3].Value = qtaInArrivo;
                            newWorksheet.Cells[rowIndex, 4].Value = ""; // Data di arrivo
                            newWorksheet.Cells[rowIndex, 5].Value = ""; // Codice fornitore
                            newWorksheet.Cells[rowIndex, 6].Value = ""; // Codice cliente

                            rowIndex++;
                        }


                        newWorkbook.SaveCsv(ConfigurationManager.AppSettings["OutputPathTipo2E3"] + "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", CsvType.SemicolonDelimited);
                        ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                        InviaEmail("MIA - Fine aggiornamento giacenze COPRE", "Il processo di aggiornamento giacenze COPRE è terminato con successo.");
                    }

                }

                //creo il file Items e il file Stock come il tracciato di Miele. Stock ogni ora, Items una volta al giorno
                // Il tipo 4 confronta i livelli del web service di GRE con quelli di COPRE.
                // Il tipo 5 è identico al tipo 3 e serve a generare lo stock
                if (tipo.Equals("4") || tipo.Equals("5"))
                {
                    //creo il file Items una volta al giorno
                    if (tipo.Equals("4"))
                    {
                        // Carica i dati dal database
                        string selQry = @"WITH CTE_Latest AS (
    SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
           cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
           flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
           flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
           flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
           prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
           ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
           ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
           ultima_presenza_nel_file_copre,
           ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
    FROM bellanti_match_articoli_gcc 
    WHERE cod_bellanti IS NULL 
      AND LEN(cod_copre) > 0 
      AND DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
      AND ordinabile_copre = 's'
      AND LEN(ean) <= 13
      AND ean IS NOT NULL
      AND LTRIM(RTRIM(ean)) <> ''
      AND prz_acquisto_copre > 0
),
CTE_FirstBrand AS (
    SELECT cod_brand, MIN(des_brand) AS des_brand
    FROM CTE_Latest
    GROUP BY cod_brand
),
CTE_Filtered AS (
    SELECT c.*
    FROM CTE_Latest c
    JOIN CTE_FirstBrand fb
    ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
)
SELECT *
FROM CTE_Filtered
WHERE rn = 1;";

                        DataTable dt = new DataTable();
                        DBAccess.ReadDataThroughAdapter(selQry, dt);

                        // Carica i dati dal file Excel con le descrizioni dei livelli
                        var workbook = new ExcelFile();
                        workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                        var worksheet = workbook.Worksheets[0];
                        var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                            .ToDictionary(
                                row => row.Cells[0].Value.ToString(),
                                row => row.Cells[1].Value.ToString()
                            );

                        // Crea un nuovo file CSV per i risultati
                        string outputPath = ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
                        using (StreamWriter sw = new StreamWriter(outputPath))
                        {
                            // Aggiungi le intestazioni
                            string[] headers = new string[] {
                "Codice_Prodotto", "Codice_EAN", "Prezzo_Acquisto", "Prezzo_Vendita", "Prezzo_Listino",
                "Spese_Consegna", "Spese_Consegna_Incluse", "Attivo", "Codice_Fornitore", "Codice_Cliente",
                "Descrizione_Breve", "Descrizione_Estesa", "Data_Lancio", "Classificazione", "Descrizione_Classificazione",
                "Codice_Marchio", "Descrizione_Marchio", "Iva", "Note", "Link_Scheda_Energetica",
                "Link_Video", "Data_Fuori_Produzione", "Raee", "Raee Annegata", "Siae", "Siae Annegata",
                "Codice_L1_Copre", "Descrizione_L1_Copre", "Codice_L2_Copre", "Descrizione_L2_Copre",
                "Codice_L3_Copre", "Descrizione_L3_Copre", "Codice_L4_Copre", "Descrizione_L4_Copre",
                "Codice_L1_WS", "Descrizione_L1_WS", "Codice_L2_WS", "Descrizione_L2_WS",
                "Codice_L3_WS", "Descrizione_L3_WS", "Codice_L4_WS", "Descrizione_L4_WS"
            };
                            sw.WriteLine(string.Join(";", headers));

                            int count = 0;
                            // Aggiungi i dati
                            foreach (DataRow row in dt.Rows)
                            {
                                string codCopre = row["cod_copre"].ToString();
                                string catEdielCopre = row["cat_ediel_copre"].ToString();
                                string desArtCopre = row["des_art_copre"].ToString();
                                string modello = ConvertToString(row["modello"]);
                                string codiceMarchio = ConvertToString(row["cod_brand"]);
                                string descrizioneMarchio = ConvertToString(row["des_brand"]);
                                string aliquota = ConvertToString(row["cod_iva"]);
                                string ean = ConvertToString(row["ean"]);
                                string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                                string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);

                                // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                                string codiceL1, descrizioneL1;
                                string codiceL2, descrizioneL2;
                                string codiceL3, descrizioneL3;
                                string codiceL4, descrizioneL4;

                                // Livello 1
                                GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                                // Livello 2
                                if (catEdielCopre.Length >= 4)
                                {
                                    GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                                }
                                else
                                {
                                    codiceL2 = "XX";
                                    descrizioneL2 = "Livello Generico";
                                }

                                // Livello 3
                                if (catEdielCopre.Length >= 6)
                                {
                                    GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                                }
                                else
                                {
                                    codiceL3 = "XX";
                                    descrizioneL3 = "Livello Generico";
                                }

                                // Livello 4
                                if (catEdielCopre.Length == 8)
                                {
                                    GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                                }
                                else
                                {
                                    codiceL4 = "XX";
                                    descrizioneL4 = "Livello Generico";
                                }

                                // Chiamata al web service per ottenere i livelli
                                string webServiceResponse = GetProductDataFromWebService(ean);
                                var (codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                                     codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS) = ParseWebServiceResponse(webServiceResponse);

                                // Scrivi i dati nella nuova riga del file CSV
                                string[] data = new string[] {
                    codCopre, ean, prezzoAcquisto, prezzoVendita, prezzoVendita,
                    "0", "0", "1", "", "",
                    desArtCopre, "", "", codiceL2, descrizioneL2,
                    codiceMarchio, descrizioneMarchio, aliquota, "", "",
                    "", "", "0", "0", "0", "0",
                    codiceL1, descrizioneL1, codiceL2, descrizioneL2,
                    codiceL3, descrizioneL3, codiceL4, descrizioneL4,
                    codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                    codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS
                };
                                sw.WriteLine(string.Join(";", data.Select(d => $"\"{d}\"")));
                                count++;
                            }
                        }

                        //ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                        Console.WriteLine("File CSV generato con successo: " + outputPath);
                    }

                    //creo il file Stock ogni ora
                    if (tipo.Equals("5"))
                    {
                        // Carica i dati dal database
                        string selQry = @"WITH CTE_Latest AS (
    SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
           cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
           flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
           flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
           flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
           prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
           ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
           ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
           ultima_presenza_nel_file_copre,
           ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
    FROM bellanti_match_articoli_gcc 
    WHERE cod_bellanti IS NULL 
      AND LEN(cod_copre) > 0 
      AND DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
      AND ordinabile_copre = 's'
      AND LEN(ean) <= 13
      AND ean IS NOT NULL
      AND LTRIM(RTRIM(ean)) <> ''
      AND prz_acquisto_copre > 0
),
CTE_FirstBrand AS (
    SELECT cod_brand, MIN(des_brand) AS des_brand
    FROM CTE_Latest
    GROUP BY cod_brand
),
CTE_Filtered AS (
    SELECT c.*
    FROM CTE_Latest c
    JOIN CTE_FirstBrand fb
    ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
)
SELECT *
FROM CTE_Filtered
WHERE rn = 1;";
                        DataTable dt = new DataTable();
                        DBAccess.ReadDataThroughAdapter(selQry, dt);

                        // Carica i dati dal file Excel con le descrizioni dei livelli
                        var workbook = new ExcelFile();
                        workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                        var worksheet = workbook.Worksheets[0];
                        var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                            .ToDictionary(
                                row => row.Cells[0].Value.ToString(),
                                row => row.Cells[1].Value.ToString()
                            );

                        // Crea un nuovo file Excel per i risultati
                        var newWorkbook = new ExcelFile();
                        var newWorksheet = newWorkbook.Worksheets.Add("Results");

                        // Aggiungi le intestazioni
                        newWorksheet.Cells[0, 0].Value = "Codice_Prodotto";
                        newWorksheet.Cells[0, 1].Value = "Codice_EAN";
                        newWorksheet.Cells[0, 2].Value = "Quantita_Disponibile";
                        newWorksheet.Cells[0, 3].Value = "Quantita_in_Arrivo";
                        newWorksheet.Cells[0, 4].Value = "Data_Arrivo";
                        newWorksheet.Cells[0, 5].Value = "Codice_Fornitore";
                        newWorksheet.Cells[0, 6].Value = "Codice_Cliente";

                        int rowIndex = 1;

                        foreach (DataRow row in dt.Rows)
                        {
                            string codCopre = row["cod_copre"].ToString();
                            string catEdielCopre = row["cat_ediel_copre"].ToString();
                            string desArtCopre = row["des_art_copre"].ToString();
                            string modello = ConvertToString(row["modello"]);
                            string codiceMarchio = ConvertToString(row["cod_brand"]);
                            string descrizioneMarchio = ConvertToString(row["des_brand"]);
                            string aliquota = ConvertToString(row["cod_iva"]);
                            string ean = ConvertToString(row["ean"]);
                            string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                            string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);
                            string qtaInArrivo = ConvertToInt(row["ordine_fornitore_copre"]);
                            string qtaDisponibile = ConvertToDecimal(row["giacenza_copre"]);

                            // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                            string codiceL1, descrizioneL1;
                            string codiceL2, descrizioneL2;
                            string codiceL3, descrizioneL3;
                            string codiceL4, descrizioneL4;

                            // Livello 1
                            GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                            // Livello 2
                            if (catEdielCopre.Length >= 4)
                            {
                                GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                            }
                            else
                            {
                                codiceL2 = "XX";
                                descrizioneL2 = "Livello Generico";
                            }

                            // Livello 3
                            if (catEdielCopre.Length >= 6)
                            {
                                GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                            }
                            else
                            {
                                codiceL3 = "XX";
                                descrizioneL3 = "Livello Generico";
                            }

                            // Livello 4
                            if (catEdielCopre.Length == 8)
                            {
                                GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                            }
                            else
                            {
                                codiceL4 = "XX";
                                descrizioneL4 = "Livello Generico";
                            }

                            // Scrivi i dati nella nuova riga del file Excel
                            newWorksheet.Cells[rowIndex, 0].Value = codCopre;
                            newWorksheet.Cells[rowIndex, 1].Value = ean;
                            newWorksheet.Cells[rowIndex, 2].Value = qtaDisponibile;
                            newWorksheet.Cells[rowIndex, 3].Value = qtaInArrivo;
                            newWorksheet.Cells[rowIndex, 4].Value = ""; // Data di arrivo
                            newWorksheet.Cells[rowIndex, 5].Value = ""; // Codice fornitore
                            newWorksheet.Cells[rowIndex, 6].Value = ""; // Codice cliente

                            rowIndex++;
                        }


                        newWorkbook.SaveCsv(ConfigurationManager.AppSettings["OutputPathTipo2E3"] + "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", CsvType.SemicolonDelimited);
                        //ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Stock_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                    }

                }

                // Il tipo 6 ottiene tutti i livelli di COPRE
                if (tipo.Equals("6"))
                {

                    // Carica i dati dal database
                    string selQry = @"WITH CTE_Latest AS (
    SELECT cod_gcc, cat_ediel, cod_brand, des_brand, ean, cod_fornitore, modello, descrizione, cod_iva, raee_siae, stato_prodotto, 
           cod_bellanti, trasferito, predefinito, prezzo_pubblico, prezzo_lordo, sconto_1_perc, flg_sconto_1_perc, sconto_2_perc, 
           flg_sconto_2_perc, sconto_3_perc, flg_sconto_3_perc, sconto_4_perc, flg_sconto_4_perc, canvass, flg_canvass, valorepiù, 
           flg_valorepiù, valoremeno, flg_valoremeno, sconto_fin, flg_sconto_fin, sconto_fa, flg_sconto_fa, perc_nc_nf, 
           flg_perc_nc_nf, val_nc_nf, flg_val_nc_nf, id_match, cod_copre, ultima_elaborazione, des_art_copre, giacenza_copre, 
           prz_acquisto_copre, prz_ordine_copre, prz_consigliato_copre, des_marchio_copre, marchio_gcc_copre, ordinabile_copre, 
           ean_copre, cat_ediel_copre, ultima_elaborazione_copre, griglia_copre, modello_copre, nuovi_arrivi_copre, 
           ordine_fornitore_copre, modello_lungo_copre, descrizione_lunga_copre, qta_rottura_copre, novita_fine_vita_copre, 
           ultima_presenza_nel_file_copre,
           ROW_NUMBER() OVER (PARTITION BY cod_copre ORDER BY ultima_presenza_nel_file_copre DESC) AS rn
    FROM bellanti_match_articoli_gcc 
    WHERE cod_bellanti IS NULL 
      AND LEN(cod_copre) > 0 
      AND DATEDIFF(day, ultima_presenza_nel_file_copre, GETDATE()) IN (0, 1, 2)
      AND ordinabile_copre = 's'
      AND LEN(ean) <= 13
      AND ean IS NOT NULL
      AND LTRIM(RTRIM(ean)) <> ''
      AND prz_acquisto_copre > 0
),
CTE_FirstBrand AS (
    SELECT cod_brand, MIN(des_brand) AS des_brand
    FROM CTE_Latest
    GROUP BY cod_brand
),
CTE_Filtered AS (
    SELECT c.*
    FROM CTE_Latest c
    JOIN CTE_FirstBrand fb
    ON c.cod_brand = fb.cod_brand AND c.des_brand = fb.des_brand
)
SELECT *
FROM CTE_Filtered
WHERE rn = 1;";

                    DataTable dt = new DataTable();
                    DBAccess.ReadDataThroughAdapter(selQry, dt);

                    // Carica i dati dal file Excel con le descrizioni dei livelli
                    var workbook = new ExcelFile();
                    workbook.LoadXlsx(ConfigurationManager.AppSettings["FileExcelLivelli"].ToString(), XlsxOptions.PreserveMakeCopy);
                    var worksheet = workbook.Worksheets[0];
                    var excelData = worksheet.Rows.Skip(1) // Salta l'intestazione
                        .ToDictionary(
                            row => row.Cells[0].Value.ToString(),
                            row => row.Cells[1].Value.ToString()
                        );

                    // Crea un nuovo file CSV per i risultati
                    string outputPath = ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
                    using (StreamWriter sw = new StreamWriter(outputPath))
                    {
                        // Aggiungi le intestazioni
                        string[] headers = new string[] {
                "Codice_Prodotto", "Codice_EAN", "Prezzo_Acquisto", "Prezzo_Vendita", "Prezzo_Listino",
                "Codice_L1_Copre", "Descrizione_L1_Copre", "Codice_L2_Copre", "Descrizione_L2_Copre",
                "Codice_L3_Copre", "Descrizione_L3_Copre", "Codice_L4_Copre", "Descrizione_L4_Copre",
                "Spese_Consegna", "Spese_Consegna_Incluse", "Attivo", "Codice_Fornitore", "Codice_Cliente",
                "Descrizione_Breve", "Descrizione_Estesa", "Data_Lancio", "Classificazione", "Descrizione_Classificazione",
                "Codice_Marchio", "Descrizione_Marchio", "Iva", "Note", "Link_Scheda_Energetica",
                "Link_Video", "Data_Fuori_Produzione", "Raee", "Raee Annegata", "Siae", "Siae Annegata",
                
                
            };
                        sw.WriteLine(string.Join(";", headers));

                        int count = 0;
                        // Aggiungi i dati
                        foreach (DataRow row in dt.Rows)
                        {
                            string codCopre = row["cod_copre"].ToString();
                            string catEdielCopre = row["cat_ediel_copre"].ToString();
                            string desArtCopre = row["des_art_copre"].ToString();
                            string modello = ConvertToString(row["modello"]);
                            string codiceMarchio = ConvertToString(row["cod_brand"]);
                            string descrizioneMarchio = ConvertToString(row["des_brand"]);
                            string aliquota = ConvertToString(row["cod_iva"]);
                            string ean = ConvertToString(row["ean"]);
                            string prezzoAcquisto = ConvertToDecimal(row["prz_acquisto_copre"]);
                            string prezzoVendita = ConvertToDecimal(row["prz_consigliato_copre"]);

                            // Verifica se catEdielCopre è null o vuoto e gestisce i livelli di conseguenza
                            string codiceL1, descrizioneL1;
                            string codiceL2, descrizioneL2;
                            string codiceL3, descrizioneL3;
                            string codiceL4, descrizioneL4;

                            // Livello 1
                            GetLevelDescription(catEdielCopre, 2, excelData, out codiceL1, out descrizioneL1);

                            // Livello 2
                            if (catEdielCopre.Length >= 4)
                            {
                                GetLevelDescription(catEdielCopre, 4, excelData, out codiceL2, out descrizioneL2);
                            }
                            else
                            {
                                codiceL2 = "XX";
                                descrizioneL2 = "Livello Generico";
                            }

                            // Livello 3
                            if (catEdielCopre.Length >= 6)
                            {
                                GetLevelDescription(catEdielCopre, 6, excelData, out codiceL3, out descrizioneL3);
                            }
                            else
                            {
                                codiceL3 = "XX";
                                descrizioneL3 = "Livello Generico";
                            }

                            // Livello 4
                            if (catEdielCopre.Length == 8)
                            {
                                GetLevelDescription(catEdielCopre, 8, excelData, out codiceL4, out descrizioneL4);
                            }
                            else
                            {
                                codiceL4 = "XX";
                                descrizioneL4 = "Livello Generico";
                            }                            

                            // Scrivi i dati nella nuova riga del file CSV
                            string[] data = new string[] {
                    codCopre, ean, prezzoAcquisto, prezzoVendita, prezzoVendita,
                    codiceL1, descrizioneL1, codiceL2, descrizioneL2,
                    codiceL3, descrizioneL3, codiceL4, descrizioneL4,
                    "0", "0", "1", "", "",
                    desArtCopre, "", "", codiceL2, descrizioneL2,
                    codiceMarchio, descrizioneMarchio, aliquota, "", "",
                    "", "", "0", "0", "0", "0"
                };
                            sw.WriteLine(string.Join(";", data.Select(d => $"\"{d}\"")));
                            count++;
                        }
                    }

                    //ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                    Console.WriteLine("File CSV generato con successo: " + outputPath);
                }

                //Il tipo 7 ottiene i livelli di SMEG a partire dal suo file excel
                if (tipo.Equals("7"))
                {
                    // Carica i dati dal foglio excel

                    string percorso = ConfigurationManager.AppSettings["PathSMEG"].ToString();
                    DataTable dt = new DataTable();

                    // Aggiungere le colonne specifiche alla DataTable
                    dt.Columns.Add("Codice_Prodotto");
                    dt.Columns.Add("Codice_EAN");
                    dt.Columns.Add("Prezzo_Acqusito");
                    dt.Columns.Add("Prezzo_Vendita");
                    dt.Columns.Add("Prezzo_Listino");
                    dt.Columns.Add("Spese_Consegna");
                    dt.Columns.Add("Spese_Consegna_Incluse");
                    dt.Columns.Add("Attivo");
                    dt.Columns.Add("Codice_Fornitore");
                    dt.Columns.Add("Codice_Cliente");
                    dt.Columns.Add("Descrizione_Breve");
                    dt.Columns.Add("Descrizione_Estesa");
                    dt.Columns.Add("Data_Lancio");
                    dt.Columns.Add("Classificazione");
                    dt.Columns.Add("Descrizione_Classificazione");
                    dt.Columns.Add("Codice_Marchio");
                    dt.Columns.Add("Descrizione_Marchio");
                    dt.Columns.Add("Iva");
                    dt.Columns.Add("Note");
                    dt.Columns.Add("Link_Scheda_Energetica");
                    dt.Columns.Add("Link_Video");
                    dt.Columns.Add("Data_Fuori_Produzione");
                    dt.Columns.Add("Raee");
                    dt.Columns.Add("Raee Annegata");
                    dt.Columns.Add("Siae");
                    dt.Columns.Add("Siae Annegata");

                    // Leggere il file di testo
                    using (StreamReader sr = new StreamReader(percorso))
                    {
                        // Leggere le righe del file di testo e aggiungere i dati alla DataTable
                        while (!sr.EndOfStream)
                        {
                            string[] fields = sr.ReadLine().Split(';'); // Cambia il delimitatore se necessario
                            DataRow row = dt.NewRow();
                            for (int i = 0; i < fields.Length-1; i++)
                            {
                                row[i] = fields[i].Trim();
                            }
                            dt.Rows.Add(row);
                        }
                    }

                    if (dt.Rows.Count > 0)
                    {
                        // Crea un nuovo file CSV per i risultati
                        string outputPath = ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_SMEG_con_livelli_impresa_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
                        using (StreamWriter sw = new StreamWriter(outputPath))
                        {
                            // Aggiungi le intestazioni
                            string[] headers = new string[] {
                "Codice_Prodotto", "Codice_EAN", "Prezzo_Acquisto", "Prezzo_Vendita", "Prezzo_Listino",
                "Spese_Consegna", "Spese_Consegna_Incluse", "Attivo", "Codice_Fornitore", "Codice_Cliente",
                "Descrizione_Breve", "Descrizione_Estesa", "Data_Lancio", "Classificazione", "Descrizione_Classificazione",
                "Codice_Marchio", "Descrizione_Marchio", "Iva", "Note", "Link_Scheda_Energetica",
                "Link_Video", "Data_Fuori_Produzione", "Raee", "Raee Annegata", "Siae", "Siae Annegata",
                "Codice_L1_Copre", "Descrizione_L1_Copre", "Codice_L2_Copre", "Descrizione_L2_Copre",
                "Codice_L3_Copre", "Descrizione_L3_Copre", "Codice_L4_Copre", "Descrizione_L4_Copre",
                "Codice_L1_WS", "Descrizione_L1_WS", "Codice_L2_WS", "Descrizione_L2_WS",
                "Codice_L3_WS", "Descrizione_L3_WS", "Codice_L4_WS", "Descrizione_L4_WS"
            };
                            sw.WriteLine(string.Join(";", headers));

                            int count = 0;
                            foreach (DataRow row in dt.Rows)
                            {
                                string ean = row[1].ToString();
                                // Chiamata al web service per ottenere i livelli
                                string webServiceResponse = GetProductDataFromWebService(ean);
                                var (codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                                     codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS) = ParseWebServiceResponse(webServiceResponse);

                                bool notFound = false;
                                if (!codiceL1_WS.Equals("ECAT NOT FOUND"))
                                {
                                    bool ok = true;


                                    string codCopre = row["Codice_Prodotto"].ToString();
                                    string desArtCopre = row["Descrizione_Breve"].ToString();
                                    string codiceMarchio = ConvertToString(row["Codice_Marchio"]);
                                    string descrizioneMarchio = ConvertToString(row["Descrizione_Marchio"]);
                                    string aliquota = ConvertToString(row["Iva"]);
                                    string prezzoAcquisto = ConvertToDecimal(row["Prezzo_Acqusito"]);
                                    string prezzoVendita = ConvertToDecimal(row["Prezzo_Vendita"]);

                                    // Scrivi i dati nella nuova riga del file CSV
                                    string[] data = new string[] {
                    codCopre, ean, prezzoAcquisto, prezzoVendita, prezzoVendita,
                    "0", "0", "1", "", "",
                    desArtCopre, "", "", "", "",
                    codiceMarchio, descrizioneMarchio, aliquota, "", "",
                    "", "", "0", "0", "0", "0",
                    "", "", "", "",
                    "", "", row["Classificazione"].ToString(), row["Descrizione_Classificazione"].ToString(),
                    codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                    codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS
                };
                                    sw.WriteLine(string.Join(";", data.Select(d => $"\"{d}\"")));
                                    count++;

                                    
                                }
                                else
                                {
                                    notFound = true;
                                }
                                

                            }

                            //ftp.UploadFile(ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", "Items_Copre_" + DateTime.Now.ToString("yyyyMMdd") + ".csv", ConfigurationManager.AppSettings["PathFolderCopre"].ToString());

                            Console.WriteLine("File CSV generato con successo: " + outputPath);


                        }
                        
                    }

                    //verifichiamo che i codici del livello 4 di gre (web service) e della nostra categoria esposizione combaciano
                    if (dt.Rows.Count > 0)
                    {
                        

                        // Crea un nuovo file Excel
                        var workbook = new ExcelFile();
                        var worksheet = workbook.Worksheets.Add("Esposizione e Livelli");

                        // Aggiungi le intestazioni
                        worksheet.Cells[0, 0].Value = "Codice_Prodotto";
                        worksheet.Cells[0, 1].Value = "Descrizione_Breve";
                        worksheet.Cells[0, 2].Value = "Codice_L4_WS";
                        worksheet.Cells[0, 3].Value = "Descrizione_L4_WS";
                        worksheet.Cells[0, 4].Value = "Codice_Cat_Esposizione";
                        worksheet.Cells[0, 5].Value = "Descrizione_Cat_Esposizione";

                        int rowIndex = 1;
                        foreach (DataRow row in dt.Rows)
                        {
                            string ean = row["Codice_EAN"].ToString();
                            string webServiceResponse = GetProductDataFromWebService(ean);

                            // Parse the WebService response
                            var (codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                                 codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS) = ParseWebServiceResponse(webServiceResponse);

                            string selQry = "select codice, descrizione from cat_esposizione where codice = '" + codiceL4_WS.Replace("'","''") + "'";
                            DataTable dtCatEsp = new();

                            DBAccess.ReadDataThroughAdapter(selQry, dtCatEsp);
                            if (dtCatEsp.Rows.Count > 0)
                            {
                                var dr = dtCatEsp.Rows[0];
                                var categoriaEsposizione = dr["codice"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["codice"].ToString()) ? dr["codice"].ToString() : "" : "";
                                var descCategoriaEsposizione = dr["descrizione"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["descrizione"].ToString()) ? dr["descrizione"].ToString() : "" : "";

                                // Verifica se il codiceL4_WS esiste nella cat_esposizione
                                if (categoriaEsposizione.Equals(codiceL4_WS))
                                {
                                        // Scrivi i dati nel file Excel
                                        worksheet.Cells[rowIndex, 0].Value = row["Codice_Prodotto"].ToString();
                                        worksheet.Cells[rowIndex, 1].Value = row["Descrizione_Breve"].ToString();
                                        worksheet.Cells[rowIndex, 2].Value = codiceL4_WS;
                                        worksheet.Cells[rowIndex, 3].Value = descrizioneL4_WS;
                                        worksheet.Cells[rowIndex, 4].Value = codiceL4_WS;
                                        worksheet.Cells[rowIndex, 5].Value = descCategoriaEsposizione;
                                        rowIndex++;                                    
                                }
                            }
                            
                        }

                        // Salva il file Excel
                        string excelOutputPath = ConfigurationManager.AppSettings["OutputPathTipo2E3"].ToString() + "Esposizione_Livelli_WS_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                        workbook.SaveXlsx(excelOutputPath);

                        Console.WriteLine("File Excel generato con successo: " + excelOutputPath);

                        //eventuale insert che potrebbe essermi d'aiuto (probabilmente però va fatta una update perchè i livelli del fornitore esterno ci sono già, ma vanno aggiornati con i nostri livelli 4)
                        //insert into t_EXSL_ExternalSupplierLevelTranscoding(exsl_exsu_code, exsl_supplierlevelcode, exsl_lev4_code, exsl_supplierleveldescription,
                        //exsl_userins, exsl_dateupd) values('#SM', 'A9', '9750', 'Lavastoviglie2', null, null)
                    }
                }

                //se il tipo è uguale a 8 faccio una importazione massiva dei livelli del fornitore esterno SMEG
                if (tipo.Equals("8"))
                {
                    // Carica i dati dal foglio excel

                    string percorso = ConfigurationManager.AppSettings["PathSMEG"].ToString();
                    DataTable dt = new DataTable();

                    // Aggiungere le colonne specifiche alla DataTable
                    dt.Columns.Add("Codice_Prodotto");
                    dt.Columns.Add("Codice_EAN");
                    dt.Columns.Add("Prezzo_Acqusito");
                    dt.Columns.Add("Prezzo_Vendita");
                    dt.Columns.Add("Prezzo_Listino");
                    dt.Columns.Add("Spese_Consegna");
                    dt.Columns.Add("Spese_Consegna_Incluse");
                    dt.Columns.Add("Attivo");
                    dt.Columns.Add("Codice_Fornitore");
                    dt.Columns.Add("Codice_Cliente");
                    dt.Columns.Add("Descrizione_Breve");
                    dt.Columns.Add("Descrizione_Estesa");
                    dt.Columns.Add("Data_Lancio");
                    dt.Columns.Add("Classificazione");
                    dt.Columns.Add("Descrizione_Classificazione");
                    dt.Columns.Add("Codice_Marchio");
                    dt.Columns.Add("Descrizione_Marchio");
                    dt.Columns.Add("Iva");
                    dt.Columns.Add("Note");
                    dt.Columns.Add("Link_Scheda_Energetica");
                    dt.Columns.Add("Link_Video");
                    dt.Columns.Add("Data_Fuori_Produzione");
                    dt.Columns.Add("Raee");
                    dt.Columns.Add("Raee Annegata");
                    dt.Columns.Add("Siae");
                    dt.Columns.Add("Siae Annegata");

                    // Leggere il file di testo
                    using (StreamReader sr = new StreamReader(percorso))
                    {
                        // Leggere le righe del file di testo e aggiungere i dati alla DataTable
                        while (!sr.EndOfStream)
                        {
                            string[] fields = sr.ReadLine().Split(';'); // Cambia il delimitatore se necessario
                            DataRow row = dt.NewRow();
                            for (int i = 0; i < fields.Length - 1; i++)
                            {
                                row[i] = fields[i].Trim();
                            }
                            dt.Rows.Add(row);
                        }
                    }

                    if (dt.Rows.Count > 0)
                    {
                        int rowIndex = 1;
                        foreach (DataRow row in dt.Rows)
                        {
                            string ean = row["Codice_EAN"].ToString();
                            string webServiceResponse = GetProductDataFromWebService(ean);

                            // Parse the WebService response
                            var (codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                                 codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS) = ParseWebServiceResponse(webServiceResponse);

                            string selQry = "select codice, descrizione from cat_esposizione where codice = '" + codiceL4_WS.Replace("'", "''") + "'";
                            DataTable dtCatEsp = new();

                            DBAccess.ReadDataThroughAdapter(selQry, dtCatEsp);
                            if (dtCatEsp.Rows.Count > 0)
                            {
                                var dr = dtCatEsp.Rows[0];
                                var exsl_supplierlevelcode = row["classificazione"] != DBNull.Value ? !String.IsNullOrWhiteSpace(row["classificazione"].ToString()) ? row["classificazione"].ToString() : "" : "";
                                var categoriaEsposizione = dr["codice"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["codice"].ToString()) ? dr["codice"].ToString() : "" : "";
                                var descCategoriaEsposizione = dr["descrizione"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["descrizione"].ToString()) ? dr["descrizione"].ToString() : "" : "";

                                // Verifica se il codiceL4_WS esiste nella cat_esposizione
                                if (categoriaEsposizione.Equals(codiceL4_WS))
                                {
                                    

                                    string updQry = "USE [MiaConsole_SIEO_Bellanti] update [t_EXSL_ExternalSupplierLevelTranscoding] set exsl_lev4_code = '" + categoriaEsposizione.Replace("'", "''") + "' where exsl_supplierlevelcode = '" + exsl_supplierlevelcode.Replace("'", "''") + "' and exsl_exsu_code = '#SM';";
                                    DBAccessSqlServer.ExecuteQuery(updQry);
                                }
                            }
                        }
                    }
                }

                //se il tipo è uguale a 9 faccio una importazione massiva dei livelli del fornitore esterno COPRE
                if (tipo.Equals("9"))
                {
                    // Carica i dati dal foglio excel

                    string percorso = ConfigurationManager.AppSettings["PathCOPRE"].ToString();
                    DataTable dt = new DataTable();

                    // Aggiungere le colonne specifiche alla DataTable
                    dt.Columns.Add("Codice_Prodotto");
                    dt.Columns.Add("Codice_EAN");
                    dt.Columns.Add("Prezzo_Acqusito");
                    dt.Columns.Add("Prezzo_Vendita");
                    dt.Columns.Add("Prezzo_Listino");
                    dt.Columns.Add("Spese_Consegna");
                    dt.Columns.Add("Spese_Consegna_Incluse");
                    dt.Columns.Add("Attivo");
                    dt.Columns.Add("Codice_Fornitore");
                    dt.Columns.Add("Codice_Cliente");
                    dt.Columns.Add("Descrizione_Breve");
                    dt.Columns.Add("Descrizione_Estesa");
                    dt.Columns.Add("Data_Lancio");
                    dt.Columns.Add("Classificazione");
                    dt.Columns.Add("Descrizione_Classificazione");
                    dt.Columns.Add("Codice_Marchio");
                    dt.Columns.Add("Descrizione_Marchio");
                    dt.Columns.Add("Iva");
                    dt.Columns.Add("Note");
                    dt.Columns.Add("Link_Scheda_Energetica");
                    dt.Columns.Add("Link_Video");
                    dt.Columns.Add("Data_Fuori_Produzione");
                    dt.Columns.Add("Raee");
                    dt.Columns.Add("Raee Annegata");
                    dt.Columns.Add("Siae");
                    dt.Columns.Add("Siae Annegata");

                    // Leggere il file di testo
                    using (StreamReader sr = new StreamReader(percorso))
                    {
                        // Salta la prima riga (intestazione delle colonne)
                        sr.ReadLine();

                        // Leggi i dati a partire dalla seconda riga
                        while (!sr.EndOfStream)
                        {
                            string[] fields = sr.ReadLine().Split(';'); // Cambia il delimitatore se necessario
                            DataRow row = dt.NewRow();
                            for (int i = 0; i < fields.Length; i++) // Assicurati di non ignorare l'ultimo campo
                            {
                                row[i] = fields[i].Replace("\"", "").Trim();
                            }
                            dt.Rows.Add(row);
                        }
                    }

                    if (dt.Rows.Count > 0)
                    {
                        int rowIndex = 1;
                        foreach (DataRow row in dt.Rows)
                        {
                            string ean = row["Codice_EAN"].ToString();
                            string webServiceResponse = GetProductDataFromWebService(ean);

                            // Parse the WebService response
                            var (codiceL1_WS, descrizioneL1_WS, codiceL2_WS, descrizioneL2_WS,
                                 codiceL3_WS, descrizioneL3_WS, codiceL4_WS, descrizioneL4_WS) = ParseWebServiceResponse(webServiceResponse);

                            string selQry = "select codice, descrizione from cat_esposizione where codice = '" + codiceL4_WS.Replace("'", "''") + "'";
                            DataTable dtCatEsp = new();

                            DBAccess.ReadDataThroughAdapter(selQry, dtCatEsp);
                            if (dtCatEsp.Rows.Count > 0)
                            {
                                var dr = dtCatEsp.Rows[0];
                                var exsl_supplierlevelcode = row["classificazione"] != DBNull.Value ? !String.IsNullOrWhiteSpace(row["classificazione"].ToString()) ? row["classificazione"].ToString() : "" : "";
                                var categoriaEsposizione = dr["codice"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["codice"].ToString()) ? dr["codice"].ToString() : "" : "";
                                var descCategoriaEsposizione = dr["descrizione"] != DBNull.Value ? !String.IsNullOrWhiteSpace(dr["descrizione"].ToString()) ? dr["descrizione"].ToString() : "" : "";

                                // Verifica se il codiceL4_WS esiste nella cat_esposizione
                                if (categoriaEsposizione.Equals(codiceL4_WS))
                                {
                                    //string checkDbQry = "SELECT DB_NAME() as CurrentDatabase;";
                                    //DataTable dtCurrentDb = new DataTable();
                                    //DBAccessSqlServer.ReadDataThroughAdapter(checkDbQry, dtCurrentDb);

                                    //if (dtCurrentDb.Rows.Count > 0)
                                    //{
                                    //    string currentDatabase = dtCurrentDb.Rows[0]["CurrentDatabase"].ToString();
                                    //    Console.WriteLine("Database corrente: " + currentDatabase);
                                    //}

                                    string updQry = "UPDATE [MiaConsole_SIEO_Bellanti].[dbo].[t_EXSL_ExternalSupplierLevelTranscoding] " +
                "SET exsl_lev4_code = '" + categoriaEsposizione.Replace("'", "''") + "' " +
                "WHERE exsl_supplierlevelcode = '" + exsl_supplierlevelcode.Replace("'", "''") + "' " +
                "AND exsl_exsu_code = '#CO';";
                                    DBAccessSqlServer.ExecuteQuery(updQry);
                                }
                            }
                        }
                    }
                }



            }
            catch (Exception ex) 
            {
                _= new  LogWriter($"Errore durante l'esecuzione del programma: {ex.Message}");
            }
        }



        private void GetLevelDescription(string code, int length, Dictionary<string, string> excelData, out string codice, out string descrizione)
        {
            if (string.IsNullOrEmpty(code) || code.Length < length)
            {
                codice = "XX";
                descrizione = "Livello Generico";
            }
            else
            {
                string key = code.Substring(0, length);
                if (excelData.TryGetValue(key, out descrizione))
                {
                    codice = key;
                }
                else
                {
                    codice = "XX";
                    descrizione = "Livello Generico";
                }
            }
        }

        private string ConvertToString(object value)
        {
            return (value == DBNull.Value || value == null) ? "" : value.ToString();
        }

        private string ConvertToDecimal(object value)
        {
            if (value == DBNull.Value || value == null)
            {
                return "0.00"; // o un valore di default appropriato
            }

            string stringValue = value.ToString();
            if (string.IsNullOrWhiteSpace(stringValue))
            {
                return "0.00"; // o un valore di default appropriato
            }

            // Rimuovi eventuali spazi vuoti iniziali e finali dal valore stringa
            stringValue = stringValue.Trim();

            // Sostituisci la virgola con il punto per gestire il formato decimale
            stringValue = stringValue.Replace(',', '.');

            if (decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return result.ToString("0.00", CultureInfo.InvariantCulture); // Formatta il decimal con due cifre decimali
            }
            else
            {
                return "0.00"; // o un valore di default appropriato
            }
        }

        private string ConvertToInt(object value)
        {
            if (value == DBNull.Value || value == null)
            {
                return "0.00"; // o un valore di default appropriato
            }

            string stringValue = value.ToString();
            if (string.IsNullOrWhiteSpace(stringValue))
            {
                return "0.00"; // o un valore di default appropriato
            }

            // Rimuovi eventuali spazi vuoti iniziali e finali dal valore stringa
            stringValue = stringValue.Trim();

            // Sostituisci la virgola con il punto per gestire il formato decimale
            stringValue = stringValue.Replace(',', '.');

            if (decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return ((int)result).ToString("0", CultureInfo.InvariantCulture);
            }
            else
            {
                return "0"; // o un valore di default appropriato
            }
        }

        public string GetProductDataFromWebService(string ecat)
        {
            string url = $"http://webservice.grespa.com:8890/webservice/?socio=BELLANTI&ecat={ecat}";

            using (WebClient client = new WebClient())
            {
                return client.DownloadString(url);
            }
        }

        public (string codiceL1, string descrizioneL1, string codiceL2, string descrizioneL2,
        string codiceL3, string descrizioneL3, string codiceL4, string descrizioneL4) ParseWebServiceResponse(string response)
        {
            var data = response.Split('|');

            if (data.Length < 9)
            {
                return ("ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND", "ECAT NOT FOUND");
            }

            return (data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8]);
        }

        // Funzione per inviare email
        void InviaEmail(string oggetto, string messaggio)
        {
            try
            {
                MailMessage msg = new MailMessage();
                msg.To.Add(ConfigurationManager.AppSettings["destinatari"].ToString());
                string pwd = ConfigurationManager.AppSettings["pwd_smtp"];
                // se vogliamo usare uno pseudonimo per il mittente, allora scriviamo new MailAddress("indirizzo","pseudonimo");
                msg.From = new MailAddress(ConfigurationManager.AppSettings["mittente"].ToString());
                msg.Subject = oggetto;
                msg.Priority = MailPriority.Normal;

                msg.IsBodyHtml = true;
                //msg.Attachments.Add(new Attachment(String.Format(System.IO.Directory.GetCurrentDirectory() + "\\log\\log_{0}.txt", DateTime.Today.ToString("yyyyMMdd"))));
                msg.Body = messaggio;
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["user_smtp"].ToString(), pwd);
                smtpClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
                smtpClient.Host = ConfigurationManager.AppSettings["smtp"];
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                try
                {
                    smtpClient.Send(msg);
                    //new LogWriter("Mail inviata correttamente");
                }
                catch (Exception ex)
                {
                    new LogWriter("Eccezione: \n" + ex.ToString());
                }
                smtpClient.Dispose();
                msg.Dispose();
            }
            catch (Exception ex)
            {
                new LogWriter("Eccezione: \n" + ex.ToString());
            }
        }

    }
}
