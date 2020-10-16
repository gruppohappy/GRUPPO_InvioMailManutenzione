using FastReport;
using FastReport.Utils;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace InvioMailManutenzione
{
    class Program
    {
        private static SqlConnection cnDb = new SqlConnection();
        private static SqlCommand cmd = new SqlCommand();
        private static DataTable tabManutenzione = new DataTable();
        private static DataTable tabRicambi = new DataTable();
        private static DataTable tabControlli = new DataTable();
        private static DataTable tabInterventiTemporanei = new DataTable();
        private static string dataNotificaAggiornata = "";

        static void Main(string[] args)
        {
            Console.WriteLine($"Avvio programma - {DateTime.Now}");
            WriteToLog($"{DateTime.Now} - Avvio programma");
            // Recupero le stringhe di connessione
            string cnString = ConfigurationManager.AppSettings["cnStringSql"].ToString();
            // Recupero la path del report
            //string reportPath = String.Concat(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), @"\REPORT_INTERVENTO.frx");
            string reportInterventoPath = ConfigurationManager.AppSettings["reportInterventoPath"].ToString();
            string reportRicambiPath = ConfigurationManager.AppSettings["reportRicambiPath"].ToString();
            // Recupero mail manutenzione
            string mailManutenzione = ConfigurationManager.AppSettings["mailManutenzione"].ToString();
            // Imposto la routine per entrambi i database (Magic ed Esperia): discrimino la destinazione in base alla connection string definita nell'app.config

            // Tento la connessione al database                
            if (CheckConnectionToDB(cnString))
            {
                /* Se true ho testato la connessione al db e procedo  */

                // Nuova implementazione: controllo presenza interventi temporanei ed inserimento
                InsertInterventoTemporaneo(cnString, reportInterventoPath, reportRicambiPath);

                // Leggo dal database se ci sono manutenzioni attive con data notifica ad oggi
                ReadFromDB(cnString);
                if (tabManutenzione.Rows.Count > 0)
                {
                    foreach (DataRow rowManutenzione in tabManutenzione.Rows)
                    {
                        if (rowManutenzione["EMAIL_RESPONSABILE"].ToString() != "")
                        {
                            
                            dataNotificaAggiornata = UpdateDataNotifica(cnString, rowManutenzione["ID_MANUTENZIONE"].ToString(), rowManutenzione["COMANDO"].ToString());
                            if (dataNotificaAggiornata != "")
                            {
                                Console.WriteLine($"Update data notifica per manutenzione #{rowManutenzione["ID_MANUTENZIONE"]} eseguito correttamente. - {DateTime.Now}");
                                WriteToLog($"{DateTime.Now} - Update data notifica per manutenzione #{rowManutenzione["ID_MANUTENZIONE"]} eseguito correttamente.");
                                // Dopo aggiornamento data notifica, spedisco invito al calendar
                                if (InvioCalendar(rowManutenzione["EMAIL_RESPONSABILE"].ToString(), rowManutenzione["TITOLO_MANUTENZIONE"].ToString(), rowManutenzione["NOME_MACCHINA"].ToString(), rowManutenzione["NOME_PERIODICITA"].ToString(), rowManutenzione["DESCRIZIONE_LAVORI"].ToString(), dataNotificaAggiornata) == true && InvioCalendar(mailManutenzione, rowManutenzione["TITOLO_MANUTENZIONE"].ToString(), rowManutenzione["NOME_MACCHINA"].ToString(), rowManutenzione["NOME_PERIODICITA"].ToString(), rowManutenzione["DESCRIZIONE_LAVORI"].ToString(), dataNotificaAggiornata) == true)
                                {
                                    Console.WriteLine($"Invio invito calendar effettuato con successo per manutenzione #{rowManutenzione["ID_MANUTENZIONE"]}");
                                    WriteToLog($"{DateTime.Now} - Invio invito calendar effettuato con successo per manutenzione #{rowManutenzione["ID_MANUTENZIONE"]}");
                                }
                                else
                                {
                                    SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE INVITO CALENDAR", $"Invio invito Google Calendar non riuscito per manutenzione #{rowManutenzione["ID_MANUTENZIONE"].ToString()}.\n");
                                }
                                // Inserisco nella tabella [LISTA_INTERVENTI] ed invio la mail all'incaricato per avvisarlo dell'inserimento
                                int idIntervento = 0;
                                idIntervento = InsertListaInterventi(cnString, rowManutenzione["TITOLO_MANUTENZIONE"].ToString(), rowManutenzione["ID_GUASTO"].ToString(), rowManutenzione["ID_MACCHINA"].ToString(), rowManutenzione["FLAG_INTERVENTO_LINEA"].ToString(), rowManutenzione["ID_UTENTE_INCARICATO"].ToString(), rowManutenzione["DESCRIZIONE_LAVORI"].ToString(), rowManutenzione["ID_MANUTENZIONE"].ToString());
                                if (idIntervento > 0)
                                {
                                    // Inserisco nella tabella LISTA_RICAMBI_UTILIZZATI
                                    if (InsertRicambiUtilizzati(cnString, rowManutenzione["ID_MANUTENZIONE"].ToString(), idIntervento))
                                    {
                                        Console.WriteLine($"Inserimento ricambi utilizzati avvenuto correttamente per la manutenzione #{rowManutenzione["ID_MANUTENZIONE"]}.");
                                        WriteToLog($"Inserimento ricambi utilizzati avvenuto correttamente per la manutenzione #{rowManutenzione["ID_MANUTENZIONE"]}.");
                                    }
                                    else
                                    {
                                        // Mando mail a IT per avvisare del mancato inserimento nei ricambi utilizzati
                                        SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE GESTIONE MANUTENZIONE", "Inserimento ricambi utilizzati non riuscito.");
                                    }
                                    // Nuova implementazione: a fronte di un intervento aperto, recupero i suoi controlli (da LISTA_CONTROLLI_MANUTENZIONE) e li inserisco nella tabella LISTA_CONTROLLI_INTERVENTO
                                    if (InsertControlliIntervento(cnString, rowManutenzione["ID_MANUTENZIONE"].ToString(), idIntervento))
                                    {
                                        // Inserimento effettuato
                                        // Se inserito correttamente, mando mail all'incaricato allegando il report intervento 
                                        InviaMailReport(reportInterventoPath, reportRicambiPath, idIntervento, rowManutenzione["EMAIL_RESPONSABILE"].ToString(), "APERTURA INTERVENTO MANUTENZIONE", $"E' stato aperto un intervento con utente incaricato: {rowManutenzione["UTENTE_INCARICATO"]}. \nIn allegato i dati.");
                                    }
                                    else
                                    {
                                        // Mando mail a IT per avvisare del mancato inserimento dei controlli intervento
                                        SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE INSERIMENTO CONTROLLI INTERVENTO", "Inserimento controlli intervento non riuscito.");
                                    }
                                }
                                else
                                {
                                    // Nel caso un cui non viene eseguito l'insert nella lista interventi, avviso via mail 
                                    SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE GESTIONE MANUTENZIONE", "Inserimento intervento non riuscito.");
                                }
                            }
                            else
                            {
                                // Nel caso in cui l'update data notifica non viene eseguito, avviso via mail
                                SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE GESTIONE MANUTENZIONE", "Update notifica non riuscito.");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Nessuna notifica mail da inviare per la data {DateTime.Now.Date.ToString("dd/MM/yyyy")}");
                    WriteToLog($"Nessuna notifica mail da inviare per la data {DateTime.Now.Date.ToString("dd / MM / yyyy")}");
                }
            }
            else
            {
                // Connessione al database non riuscita, chiudo il programma
                Console.WriteLine($"Chiusura programma - {DateTime.Now}\n");
                WriteToLog($"Chiusura programma - {DateTime.Now}\n");
            }

            // Ho finito il processo e chiudo il programma
            Console.WriteLine($"Chiusura programma - {DateTime.Now}\n");
            WriteToLog($"Chiusura programma - {DateTime.Now}\n");
        }

        /* Metodo che controlla la presenza di interventi temporanei con data nuovo intervento ad oggi, se si viene creato un nuovo intervento*/
        private static void InsertInterventoTemporaneo(string _connectionString, string _reportInterventoPath, string _reportRicambiPath)
        {
            try
            {
                tabInterventiTemporanei.Clear();
                string today = DateTime.Now.Date.ToString("yyyy-MM-dd");
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.Connection = cnDb;
                cmd.CommandText = $"SELECT * FROM ANAGRAFICA_V_INTERVENTI_TEMPORANEI WHERE FLAG_INTERVENTO_TEMPORANEO = 1 AND DATA_NUOVO_INTERVENTO = '{today}'";
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(tabInterventiTemporanei);
                }

                if (tabInterventiTemporanei.Rows.Count > 0)
                {
                    // Ho trovato interventi temporanei
                    foreach (DataRow row in tabInterventiTemporanei.Rows)
                    {                       
                        int ID = 0;
                        // Inserisco un nuovo intervento temporaneo con stessi dati e con ID_INTERVENTO_COLLEGATO il mio ID di partenza
                        cmd.CommandText = "INSERT INTO [LISTA_INTERVENTI] (DATA_APERTURA, DATA_CHIUSURA, DATA_ELABORAZIONE, TITOLO, TEMPO_ESECUZIONE, ID_STATO, ID_GUASTO," +
                            "ID_MACCHINA, FLAG_INTERVENTO_LINEA, PRIORITA, ID_UTENTE_APERTURA, ID_UTENTE_ESECUTORE, DESCRIZIONE_GUASTO," +
                            "ID_ESITO, DESCRIZIONE_LAVORO, DESCRIZIONE_RITARDO, DATA_NUOVO_INTERVENTO, FLAG_INTERVENTO_TEMPORANEO ,ID_MANUTENZIONE_COLLEGATA, ID_INTERVENTO_TEMPORANEO_COLLEGATO) OUTPUT Inserted.ID_INTERVENTO VALUES (" +
                            $"CONVERT(DATE,'{DateTime.Now.Date.ToString("yyyy-MM-dd")}')," + // Data apertura 
                            $"NULL," +      // Data chiusura
                            $"NULL," +      // Data elaborazione
                            $"'{row["TITOLO"].ToString()}'," +      // TITOLO
                            $"NULL," +    //TEMPO ESECUZIONE
                            $"1," +       // ID STATO
                            $"'{row["ID_GUASTO"].ToString()}'," +       // ID GUASTO
                            $"'{row["ID_MACCHINA"].ToString()}'," +     // ID MACCHINA 
                            $"'{row["FLAG_INTERVENTO_LINEA"].ToString()}'," +       // FLAG INTERVENTO LINEA
                            $"'{row["PRIORITA"].ToString()}'," +        //  PRIORITA
                            $"'{row["ID_UTENTE_APERTURA"].ToString()}'," +      // ID_UTENTE_APERTURA
                            $"NULL," +        // ID UTENTE ESECUTORE
                            $"'{row["DESCRIZIONE_GUASTO"].ToString()}'," +      // DESCRIZIONE GUASTO
                            $"NULL," +        // ID ESITO
                            $"NULL," +        // DESCRIZIONE LAVORO
                            $"NULL," +        // DESCRIZIONE RITARDO 
                            $"NULL," +          // DATA NUOVO INTERVENTO
                            $"'{row["FLAG_INTERVENTO_TEMPORANEO"].ToString()}'," +  // FLAG INTERVENTO TEMPORANEO
                            $"'{row["ID_MANUTENZIONE_COLLEGATA"].ToString()}'," +       // ID MANUTENZIONE COLLEGATA
                            $"'{row["ID_INTERVENTO"].ToString()}'" +    // ID_INTERVENTO_TEMPORANEO_ASSOCIATO --> E' l'ID INTERVENTO dell'intervento di partenza.
                            ")";
                        ID = Convert.ToInt32(cmd.ExecuteScalar());
                        cnDb.Close();
                        if (ID > 0)
                        {
                            Console.WriteLine($"Creazione nuovo intervento temporaneo effettuata per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}. ID associato: {ID}");
                            WriteToLog($"Creazione nuovo intervento temporaneo effettuata per il database { (_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}. ID associato: { ID}");
                          
                            // Nuova implementazione: a fronte di un intervento aperto, recupero i suoi controlli (da LISTA_CONTROLLI_MANUTENZIONE) e li inserisco nella tabella LISTA_CONTROLLI_INTERVENTO
                            if (InsertControlliIntervento(_connectionString, row["ID_MANUTENZIONE_COLLEGATA"].ToString(), ID))
                            {
                                // Inserimento effettuato
                                // Se inserito correttamente, mando mail all'incaricato allegando il report intervento                                 
                                InviaMailReport(_reportInterventoPath, _reportRicambiPath, ID, row["EMAIL_RESPONSABILE"].ToString(), "APERTURA INTERVENTO TEMPORANEO", $"E' stato aperto un intervento temporaneo. \nIn allegato i dati.");
                            }
                            else
                            {
                                // Mando mail a IT per avvisare del mancato inserimento dei controlli intervento
                                SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE INSERIMENTO CONTROLLI INTERVENTO", "Inserimento controlli intervento non riuscito.");
                            }
                        }
                    }
                }
                else
                {
                    // Non c'è nessun intervento temporaneo
                    Console.WriteLine($"{DateTime.Now} - Nessun intervento temporaneo trovato.");
                    return;
                }
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now} - {ex.Message}");
                WriteToLog($"{DateTime.Now} - {ex.Message}");
                return;
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Metodo di invio invito Google Calendar via mail */
        private static bool InvioCalendar(string _email, string _titolo, string _nomeMacchina, string _periodicita, string _descrizione, object _dataNotifica)
        {
            try
            {
                if (_email != "")
                {
                    string oggettoCalendar = $"{_titolo} - {_nomeMacchina} - {_periodicita}";
                    string corpoCalendar = _descrizione;
                    DateTime dataEvento = Convert.ToDateTime(_dataNotifica);

                    MailMessage msg = new MailMessage();
                    SmtpClient sc = new SmtpClient("smtp.gmail.com", 587);
                    msg.From = new MailAddress("support@gruppo-happy.it", _titolo);
                    sc.Credentials = new NetworkCredential("support@gruppo-happy.it", "7t9Pe!aB");
                    sc.EnableSsl = true;

                    msg.To.Add(new MailAddress(_email, _email));
                    msg.Subject = oggettoCalendar;
                    msg.Body = corpoCalendar;

                    StringBuilder str = new StringBuilder();
                    str.AppendLine("BEGIN:VCALENDAR");
                    str.AppendLine("PRODID:-//GeO");
                    str.AppendLine("VERSION:2.0");
                    str.AppendLine("METHOD:REQUEST");
                    str.AppendLine("BEGIN:VEVENT");
                    //str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", _dataEvento));
                    //str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
                    //str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", _dataEvento));
                    str.AppendLine(string.Format("DTSTART:{0:yyyyMMdd}", dataEvento));
                    str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMdd}", DateTime.UtcNow));
                    str.AppendLine(string.Format("DTEND:{0:yyyyMMdd}", dataEvento));
                    //str.AppendLine("LOCATION: " + Direccion);
                    str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                    //str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
                    str.AppendLine(string.Format("DESCRIPTION;ENCODING=QUOTED-PRINTABLE:{0}", msg.Body));

                    str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
                    str.AppendLine(string.Format("SUMMARY;ENCODING=QUOTED-PRINTABLE:{0}", msg.Subject));
                    str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

                    str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));

                    str.AppendLine("BEGIN:VALARM");
                    str.AppendLine("TRIGGER:-PT15M");
                    str.AppendLine("ACTION:DISPLAY");
                    str.AppendLine("DESCRIPTION;ENCODING=QUOTED-PRINTABLE:Reminder");
                    str.AppendLine("END:VALARM");
                    str.AppendLine("END:VEVENT");
                    str.AppendLine("END:VCALENDAR");
                    System.Net.Mime.ContentType type = new System.Net.Mime.ContentType("text/calendar");
                    type.Parameters.Add("method", "REQUEST");
                    //type.Parameters.Add("method", "PUBLISH");
                    //type.Parameters.Add("name", "Cita.ics");
                    msg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(str.ToString(), type));
                    sc.Send(msg);
                    return true;
                }
                else
                {
                    WriteToLog($"{DateTime.Now} - Mail invito calendar non compilata.");
                    return false;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now} - Errore invio invito calendar:\n{ex.ToString()}");
                WriteToLog($"{DateTime.Now} - Errore invio invito calendar:\n{ex.ToString()}");
                return false;
            }
        }

        /* Inserimento controlli nella tabella LISTA_CONTROLLI_INTERVENTO */
        private static bool InsertControlliIntervento(string _connectionString, string _idManutenzione, int _idIntervento)
        {
            try
            {
                // Recupero i controlli associati alla manutenzione di cui è stato inserito l'intervento
                tabControlli.Clear();
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.CommandText = "INSERT INTO [LISTA_CONTROLLI_INTERVENTO] (ID_CONTROLLO, ID_INTERVENTO) " +
                    $"SELECT ID_CONTROLLO, {_idIntervento} " +
                    $"FROM [LISTA_CONTROLLI_MANUTENZIONE] WHERE ID_MANUTENZIONE = {_idManutenzione}";
                cmd.ExecuteScalar();
                cnDb.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore nell'inserimento controlli intervento:\n{ex.Message}");
                WriteToLog($"{DateTime.Now} - Errore inserimento controlli intervento:\n{ex.ToString()}");
                return false;
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Inserimento ricambi */
        private static bool InsertRicambiUtilizzati(string _connectionString, string _idManutenzione, int _idIntervento)
        {
            try
            {
                tabRicambi.Clear();
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.CommandText = "INSERT INTO [LISTA_RICAMBI_UTILIZZATI] (CODICE_RICAMBIO, QTA_UTILIZZATA, ID_INTERVENTO, SCARICATO, PROPOSTO)" +
                    $"SELECT CODICE_RICAMBIO, QTA_PROPOSTA, {_idIntervento}, 0, 1" +
                    $"FROM [LISTA_RICAMBI_MANUTENZIONE] WHERE ID_MANUTENZIONE = {_idManutenzione}";
                cmd.ExecuteScalar();
                cnDb.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore nell'inserimento ricambi utilizzati:\n{ex.Message}");
                WriteToLog($"{DateTime.Now} - Errore nell'inserimento ricambi utilizzati:\n{ex.Message}");
                return false;
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Inserisco un nuovo intervento nella tabella LISTA_INTERVENTI */
        private static int InsertListaInterventi(string _connectionString, string _titoloManutenzione, string _idGuasto, string _idMacchina, string _flagInterventoLinea, string _idUtenteIncaricato, string _descrizioneLavori, string _manutenzioneCollegata)
        {
            try
            {
                int ID = 0;
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.CommandText = "INSERT INTO [LISTA_INTERVENTI] (DATA_APERTURA, DATA_CHIUSURA, DATA_ELABORAZIONE, TITOLO, TEMPO_ESECUZIONE, ID_STATO, ID_GUASTO," +
                    "ID_MACCHINA, FLAG_INTERVENTO_LINEA, PRIORITA, ID_UTENTE_APERTURA, ID_UTENTE_ESECUTORE, DESCRIZIONE_GUASTO," +
                    "ID_ESITO, DESCRIZIONE_LAVORO, DESCRIZIONE_RITARDO, DATA_NUOVO_INTERVENTO, ID_MANUTENZIONE_COLLEGATA) OUTPUT Inserted.ID_INTERVENTO VALUES (" +
                    $"CONVERT(DATE,'{DateTime.Now.Date.ToString("yyyy-MM-dd")}')," +
                    "NULL," +
                    "NULL," +
                    $"'{_titoloManutenzione}'," +
                    "NULL," +
                    "1," +
                    $"{_idGuasto}," +
                    $"{_idMacchina}," +
                    $"'{_flagInterventoLinea}'," +
                    "1," +
                    $"{_idUtenteIncaricato}," +
                    "NULL," +
                    $"'{_descrizioneLavori}'," +
                    "NULL," +
                    "NULL," +
                    "NULL," +
                    "NULL," +
                    $"{_manutenzioneCollegata}" +
                    ")";
                ID = Convert.ToInt32(cmd.ExecuteScalar());
                if (ID > 0)
                {
                    Console.WriteLine($"Creazione nuovo intervento effettuata per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}. ID associato: {ID}");
                    WriteToLog($"Creazione nuovo intervento effettuata per il database { (_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}. ID associato: { ID}");
                    return ID;
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ATTENZIONE: Creazione intervento non riuscito per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:");
                Console.WriteLine($"{ex.Message}");
                WriteToLog($"ATTENZIONE: Creazione intervento non riuscito per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:");
                WriteToLog($"{ex.Message}");
                return 0;
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Invio la mail di notifica  */
        private static bool SendMail(string _fromMail, string _toMail, string _subject, string _body)
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    using (SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com"))
                    {
                        mail.From = new MailAddress(_fromMail);
                        mail.To.Add(_toMail);
                        mail.Subject = _subject;
                        mail.Body = _body;
                        SmtpServer.Port = 587;
                        SmtpServer.Credentials = new System.Net.NetworkCredential(_fromMail, "7t9Pe!aB");
                        SmtpServer.EnableSsl = true;
                        SmtpServer.Send(mail);
                        Console.WriteLine($"Mail inviata a {_toMail}. - {DateTime.Now}");
                        WriteToLog($"Mail inviata a {_toMail}. - {DateTime.Now}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore nell'invio della mail:\n{ex.Message}");
                WriteToLog($"Errore nell'invio della mail:\n{ex.Message} - {DateTime.Now}");
                return false;
            }
        }

        /* Lettura lista manutenzioni dal db */
        private static void ReadFromDB(string _connectionString)
        {
            try
            {
                tabManutenzione.Clear();
                string today = DateTime.Now.Date.ToString("yyyy-MM-dd");
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.Connection = cnDb;
                cmd.CommandText = $"SELECT * FROM ANAGRAFICA_V_MANUTENZIONI_ATTIVE WHERE ATTIVO = 1 AND DATA_NOTIFICA = '{today}'";
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(tabManutenzione);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore in fase di lettura dal database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:\n{ex.Message}");
                WriteToLog($"{DateTime.Now} - Errore in fase di lettura dal database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:\n{ex.Message}");
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Eseguo un update del campo DATA_NOTIFICA, posticipandolo in base alla sua periodicità */
        private static string UpdateDataNotifica(string _connectionString, string _IdManutenzione, string _comando)
        {
            try
            {
                string UpdateDataNotifica = _comando.Replace("DATA_APERTURA", "DATA_NOTIFICA");
                string dataNotificaAggiornata = "";
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cmd.CommandText = "UPDATE [GESTIONE_MANUTENZIONE].[dbo].[LISTA_MANUTENZIONI] SET [DATA_NOTIFICA] = " + UpdateDataNotifica + $" OUTPUT INSERTED.DATA_NOTIFICA WHERE ID_MANUTENZIONE = {_IdManutenzione}";
                dataNotificaAggiornata = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("yyyy-MM-dd");

                cnDb.Close();
                return dataNotificaAggiornata;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ATTENZIONE: Update data notifica per manutenzione #{_IdManutenzione} non riuscito per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:");
                Console.WriteLine($"{ex.Message}");
                WriteToLog($"{DateTime.Now} - ATTENZIONE: Update data notifica per manutenzione #{_IdManutenzione} non riuscito per il database {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}:\n{ex.Message}");
                return "";
            }
            finally
            {
                if (cnDb.State != ConnectionState.Closed)
                {
                    cnDb.Close();
                }
            }
        }

        /* Check connessione al database */
        private static bool CheckConnectionToDB(string _connectionString)
        {
            try
            {
                // Apro la connessione
                cnDb.ConnectionString = _connectionString;
                cnDb.Open();
                cnDb.Close();
                Console.WriteLine($"Connessione al server {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")} riuscita.");
                WriteToLog($"Connessione al server {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")} riuscita");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore in fase di check connessione al server {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}: {ex.Message}");
                SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", "ERRORE GESTIONE MANUTENZIONE", $"Connessione a SRVDATABASE/SQLGH non riuscita: {ex.Message}");
                WriteToLog($"{DateTime.Now} - Errore in fase di check connessione al server {(_connectionString.Contains(@"SRVDATABASE\SQLGH") ? @"SRVDATABASE\SQLGH" : @"SRVTSESP\SQLESPERIA")}: {ex.Message}");
                return false;
            }
        }

        /* Gestione log testuale */
        private static void WriteToLog(string _message)
        {
            string programPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string logFileName = String.Concat(programPath, $@"\Log\", $"Logfile_{DateTime.Now.Date.ToString("dd-MM-yyyy")}.txt");
            // Controllo che la cartella dei log esista, in caso contrario la creo
            if (!Directory.Exists(String.Concat(programPath, @"\Log")))
            {
                // Creo  la cartella Log
                Directory.CreateDirectory(String.Concat(programPath, @"\Log"));
            }

            // Controllo che il file txt di log esista, in caso contrario lo creo
            if (File.Exists(logFileName))
            {
                // File esiste, aggiungo all coda una riga di testo
                using (StreamWriter sw = File.AppendText(logFileName))
                {
                    sw.WriteLine(_message);
                }
            }
            else
            {
                // File non esistente, lo creo
                FileStream fs = File.Create(logFileName);
                fs.Close();
                using (StreamWriter sw = File.AppendText(logFileName))
                {
                    sw.WriteLine(_message);
                }
            }
        }

        /* Metodo di invio report via Mail da utilizzare in caso di creazione intervento */
        private static bool InviaMailReport(string _nomeReportIntervento, string _nomeReportRicambi, int _idIntervento, string _destinatario, string _titolo, string _corpo)
        {
            try
            {
                Config.ReportSettings.ShowProgress = false; // Disabilito finestra di progresso
                Report report1 = new Report(); // Creo nuovo oggetto di report per report intervento
                Report report2 = new Report();  // Creo nuovo oggetto di report per report ricambi
                report1.Load(_nomeReportIntervento); // Carico il report intervento
                report2.Load(_nomeReportRicambi); // Carico il report ricambi
                report1.SetParameterValue("idIntervento", _idIntervento);
                report2.SetParameterValue("idIntervento", _idIntervento);
                report1.Prepare();
                report2.Prepare();
                FastReport.Export.Pdf.PDFExport pdf = new FastReport.Export.Pdf.PDFExport(); // Esporto il PDF
                FastReport.Export.Email.EmailExport email = new FastReport.Export.Email.EmailExport(); // Esporto il report via mail
                // Impostazioni invio mail
                email.Account.Address = "support@gruppo-happy.it";
                email.Account.Name = "GESTIONE MANUTENZIONE";
                email.Account.Host = "smtp.gmail.com";
                email.Account.Port = 25;
                email.Account.UserName = "support@gruppo-happy.it";
                email.Account.Password = "7t9Pe!aB";
                email.Account.EnableSSL = true;
                // Impostazioni indirizzo mail
                email.Address = _destinatario;
                email.Subject = _titolo;
                email.MessageBody = _corpo;
                email.Export = pdf; // Setto tipo esportazione
                email.SendEmail(report1, report2); // Invio mail
                Console.WriteLine($"{DateTime.Now} - Invio report intervento e ricambi #{_idIntervento} via mail effettuato.");
                WriteToLog($"{DateTime.Now} - Invio report intervento e ricambi #{_idIntervento} via mail effettuato.");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now} - Invio report intervento e ricambi #{_idIntervento} via mail non riuscito: {ex.Message}");
                WriteToLog($"{DateTime.Now} - Invio report intervento e ricambi #{_idIntervento} via mail non riuscito: {ex.Message}");
                //SendMail("support@gruppo-happy.it", "it@gruppo-happy.it", $"Errore invio report intervento e ricambi #{_idIntervento}", $"Errore invio report intervento con ID in oggetto.\n{ex.Message}");
                return false;
            }
        }
    }
}
