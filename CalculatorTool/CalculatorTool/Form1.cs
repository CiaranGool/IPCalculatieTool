using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalculatorTool
{
    // latest version
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region String en Doubles

        //String waarden stickers verwijderen en aanbrengen
        string Oppervlakte, Aantal, Aantal_Messen, Ver_Werkuren, Ver_Middel1, Ver_Middel2, Ver_Middel3;
        string Aan_Oppervlakte, Aan_Aantal, Aan_Middel1, Aan_Werkuren;

        //Double waarden stickers verwijderen en aanbrengen
        double WaardeOppervlakte, WaardeAantal, WaardeMessen, Waarde_Ver_Werkuren, Ver_WerkurenPrijs, Ver_WaardeMiddel1, Ver_WaardeMiddel2, Ver_WaardeMiddel3;
        double Aan_WaardeOppervlakte, Aan_WaardeAantal, Aan_WaardeMiddel1, Aan_Waarde_Werkuren, Aan_WerkurenPrijs;

        //String en double waarden voor reistijd
        string Reis_Vergoeding, Reistijd, PrijsBenzine, ParkeerKosten, Afstand;
        double WaardeReiskosten, WaardeReistijd, WaardeReis_Vergoeding, WaardePrijsBenzine, WaardeParkeerKosten, WaardeAfstand;

        //String en double waarden voor loonberekening
        string Naam1, Naam2, Naam3, Naam4, Uren1, Uren2, Uren3, Uren4, LPU1, LPU2, LPU3, LPU4, TotaleWerkuren1, TotaleWerkuren2, TotaleWerkuren3, TotaleWerkuren4;
        double WaardeUren1, WaardeUren2, WaardeUren3, WaardeUren4, WaardeLPU1, WaardeLPU2, WaardeLPU3, WaardeLPU4, WaardeTotaleWerkuren1, WaardeTotaleWerkuren2, WaardeTotaleWerkuren3, WaardeTotaleWerkuren4;

        //String en double waarden voor eigenschappen opdracht
        string Type_Klus, OrderNummer, Locatie, Temperatuur, Voorbereidingstijd;
        double WaardeOrderNummer, WaardeTemperatuur, WaardeVoorbereidingstijd;


        private void btn_Hulp_Click(object sender, EventArgs e)
        {
            string Bericht = "Met behulp van de Automatiseringstool is het mogelijk om de kosten en verdiensten te berkenen, om vervolgens het winstmarge te kunnen bepalen." + "\n" + "\n" +
                  "U dient de Invoer kosten van verwijdering, Invoer kosten van aanbreng of beide volledig in te vullen om een correcte berekening uit te voeren." + "\n" + "\n" +
                  "Vervolgens moeten de eigenschappen van de klus gespecificeerd worden, en moeten de reiskosten worden ingevoerd." + "\n" + "\n" +
                  "Tot slot is het nog mogelijk om de loon van de werknemers in te voeren voor nog een duidelijker overzicht op het winstmarge." + "\n" + "\n" +"\n" + "\n" +
                  "Door op de knop berekenen te drukken is het mogelijk om een overzicht van alle gegevens te creëren in de tool zelf." + "\n" + "\n" +
                  "Door op de knop exporteren te drukken is het mogelijk om een gedetailleerd overzicht in Excel te creëren.";
            string Titel = "Uitleg voor het gebruik";

            MessageBox.Show(Bericht, Titel);
        }

        string Datum, Opmerkingen, RollenPoetspapier, ExtraBenodigdheden;
        double EindberekeningKosten, EindberekeningVerdiensten, EindberekeningWinst, WaardeRollenPoetspapier, WaardeExtraBenodigdheden;

        #endregion

        private void txt_AanBreedte_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_AanBreedte, false);
        }

        private void txt_AanHoogte_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_AanHoogte, false);
        }

        #region Invoercheck
        private void txt_Temperatuur_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Temperatuur, false);
        }

        private void txt_Voorbereidingstijd_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Voorbereidingstijd, false);
        }

        private void txt_VerwijderBreedte_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_VerwijderBreedte, false);
        }

        private void txt_VerwijderHoogte_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_VerwijderHoogte, false);
        }
        #endregion
       
        private void txt_OrderNummer_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_OrderNummer, false);
        }

        

        //Private voids van alle textboxes en comboboxes KleurCheckers 
        #region priv voids

        #region Invoercheck
        // Kleurcontrole voor invoer in de textboxen.
        private void txt_Type_Klus_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Type_Klus, true);
        }

        private void txt_Messen_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Messen, false);
        }

        private void txt_WerkurenVerwijderen_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_WerkurenVerwijderen, false);
        }

        private void txt_locatie_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_locatie, true);
        }

        private void txt_RolPapier_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_RolPapier, false);
        }

        private void txt_OverigeKosten_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_OverigeKosten, false);
        }

        private void txt_OpenbaarVervoerKosten_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_OpenbaarVervoerKosten, false);
        }

        private void txt_ParkeerKosten_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_ParkeerKosten, false);
        }

        private void txt_PrijsBenzine_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_PrijsBenzine, false);
        }

        private void txt_Afstand_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Afstand, false);
        }

        private void txt_Naam1_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Naam1, true);
        }

        private void txt_Naam2_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Naam2, true);
        }

        private void txt_Naam3_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Naam3, true);
        }

        private void txt_Naam4_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Naam4, true);
        }

        private void txt_LPU1_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_LPU1, false);
        }

        private void txt_LPU2_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_LPU2, false);
        }

        private void txt_LPU3_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_LPU3, false);
        }

        private void txt_LPU4_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_LPU4, false);
        }

        private void txt_TotaleWerkuren1_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_TotaleWerkuren1, false);
        }

        private void txt_TotaleWerkuren2_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_TotaleWerkuren2, false);
        }

        private void txt_TotaleWerkuren3_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_TotaleWerkuren3, false);
        }

        private void txt_TotaleWerkuren4_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_TotaleWerkuren4, false);
        }

        private void txt_Middel1_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Middel1, false);
        }

        private void txt_Middel2_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Middel2, false);
        }

        private void txt_Middel3_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Middel3, false);
        }

        private void txt_Aantal_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Aantal, false);
        }


        private void cmb_Prijs_Per_m2_DrawItem(object sender, DrawItemEventArgs e)
        {
            int index = e.Index >= 0 ? e.Index : 0;
            var brush = Brushes.Black;
            e.DrawBackground();
            e.Graphics.DrawString(cmb_Prijs_Per_m2.Items[index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();

            InvoerCheckCmb(cmb_Prijs_Per_m2, false);
        }


        private void txt_OppervlakteAanbrengen_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_OppervlakteAanbrengen, false);
        }

        private void txt_AantalAanbrengen_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_AantalAanbrengen, false);
        }

        private void txt_AanbrengenMiddel1_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_AanbrengenMiddel1, false);
        }

        private void txt_WerkurenAanbrengen_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_WerkurenAanbrengen, false);
        }

        private void txt_Reistijd_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Reistijd, false);
        }

        private void txt_ReisVergoeding_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_ReisVergoeding, false);
        }
        #endregion

        #region cmb Boxen
        private void cmb_PPU_Verwijderen_DrawItem(object sender, DrawItemEventArgs e)
        {
            int index = e.Index >= 0 ? e.Index : 0;
            var brush = Brushes.Black;
            e.DrawBackground();
            e.Graphics.DrawString(cmb_PPU_Verwijderen.Items[index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();

            InvoerCheckCmb(cmb_PPU_Verwijderen, false);
        }

        private void cmb_Type_Materiaal_DrawItem(object sender, DrawItemEventArgs e)
        {
            int index = e.Index >= 0 ? e.Index : 0;
            var brush = Brushes.Black;
            e.DrawBackground();
            e.Graphics.DrawString(cmb_Type_Materiaal.Items[index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();

            InvoerCheckCmb(cmb_Type_Materiaal, true);
        }

        private void cmb_PPU_Aanbrengen_DrawItem(object sender, DrawItemEventArgs e)
        {
            int index = e.Index >= 0 ? e.Index : 0;
            var brush = Brushes.Black;
            e.DrawBackground();
            e.Graphics.DrawString(cmb_PPU_Aanbrengen.Items[index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();

            InvoerCheckCmb(cmb_PPU_Aanbrengen, false);
        }

        private void cmb_AantalWerknemers_DrawItem(object sender, DrawItemEventArgs e)
        {
            int index = e.Index >= 0 ? e.Index : 0;
            var brush = Brushes.Black;
            e.DrawBackground();
            e.Graphics.DrawString(cmb_AantalWerknemers.Items[index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();

            InvoerCheckCmb(cmb_AantalWerknemers, false);
        }
        #endregion


        private void btn_ExporterenNaarExcel_Click(object sender, EventArgs e)
        {
            #region Initialisatie 
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Range oRng;

            Oppervlakte = txt_Oppervlakte.Text;
            Aantal = txt_Aantal.Text;
            Type_Klus = txt_Type_Klus.Text;
            Aantal_Messen = txt_Messen.Text;

            Aan_Oppervlakte = txt_OppervlakteAanbrengen.Text;
            Aan_Aantal = txt_AantalAanbrengen.Text;

            Loonberekening();
            EigenschappenOpdrachtberekening();
            //Open excel
            oXL = new Excel.Application();
            oXL.Visible = true;


            //Nieuw document aanmaken.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                

            //Titels stickers Verwijderen
            oSheet.Cells[3, 1] = "Artikel soort";
            oSheet.Cells[3, 5] = "M²";
            oSheet.Cells[3, 6] = "Aantal";
            oSheet.Cells[3, 7] = "Prijs";
            oSheet.Cells[3, 8] = "Totaal";
            oSheet.Cells[5, 1] = "Kosten werkzaamheden";
            oSheet.Cells[13, 1] = "Verwijder kosten totaal";
            oSheet.Cells[15, 1] = "Opbrengst werkzaamheden";
            oSheet.Cells[19, 1] = "Reisvergoeding";
            oSheet.Cells[22, 1] = "Kosten reizen";

            oSheet.Cells[19, 5] = "Reistijd";
            oSheet.Cells[19, 6] = "Personen";
            oSheet.Cells[19, 7] = "PPU";
            oSheet.Cells[19, 8] = "Totaal";


            oSheet.Cells[18, 17] = "Totale kosten";
            oSheet.Cells[19, 17] = "Totale verdiensten";
            oSheet.Cells[21, 17] = "Winstmarge";

            oSheet.Cells[22, 5] = "Parkeren";
            oSheet.Cells[22, 6] = "Benzine";
            oSheet.Cells[22, 7] = "Afstand";

            oSheet.Cells[26, 1] = "Eigenschappen Opdracht";

            //test loonberekening
            //oSheet.get_Range("W3", "W3").Value2 = Loonberekening(WaardeLPU3);
            #endregion

            #region Opmaak
            //Opmaak.
            oSheet.get_Range("A3", "V5").Font.Bold = true;
            oSheet.get_Range("A3", "D3").VerticalAlignment =
            Excel.XlVAlign.xlVAlignCenter;
    
            //leeg kader
            oSheet.get_Range("E6", "E6").HorizontalAlignment =
            Excel.XlHAlign.xlHAlignRight;

            //Font kleuren en dikte
            oSheet.get_Range("A3", "V5").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A1", "Z1").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A6", "S11").Font.Color = Color.SlateBlue;
            oSheet.get_Range("A16", "G16").Font.Color = Color.SlateBlue;
            oSheet.get_Range("H16", "H16").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A12", "H12").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A13", "V13").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A15", "V15").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("L16", "R16").Font.Color = Color.SlateBlue;
            oSheet.get_Range("A20", "H20").Font.Color = Color.SlateBlue;
            oSheet.get_Range("AA5", "AC8").Font.Color = Color.SlateBlue;
            oSheet.get_Range("A19", "H19").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A22", "G22").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A23", "A23").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("Q18", "Q21").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("W3", "AC4").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A13", "V13").Font.Bold = true;
            oSheet.get_Range("W3", "AC4").Font.Bold = true;
            oSheet.get_Range("A15", "V15").Font.Bold = true;
            oSheet.get_Range("A19", "H19").Font.Bold = true;
            oSheet.get_Range("A22", "G22").Font.Bold = true;
            oSheet.get_Range("A23", "A23").Font.Bold = true;
            oSheet.get_Range("Q18", "Q21").Font.Bold = true;


            oSheet.get_Range("A1", "Z1").Font.Bold = true;
            oSheet.get_Range("A1", "Z1").Font.Size = 16;
            oSheet.get_Range("A1", "A1").Value2 = "Stickers verwijderen";

            oSheet.get_Range("A26", "A26").Font.Bold = true;
            oSheet.get_Range("A26", "A26").Font.Size = 16;
            oSheet.get_Range("A26", "V26").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A28", "A40").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("W5", "W8").Font.Color = Color.MidnightBlue;
            oSheet.get_Range("A28", "A40").Font.Bold = true;
            oSheet.get_Range("W5", "W8").Font.Bold = true;
            oSheet.get_Range("E28", "H40").Font.Color = Color.SlateBlue;


            #endregion

            #region Verwijderen

            //Kleur achtergrond kaders
            oSheet.get_Range("H13", "H13").Interior.Color = Color.LightCoral;
            oSheet.get_Range("S13", "S13").Interior.Color = Color.LightCoral;
            oSheet.get_Range("H16", "H16").Interior.Color = Color.LightGreen;
            oSheet.get_Range("S16", "S16").Interior.Color = Color.LightGreen;
            oSheet.get_Range("H20", "H20").Interior.Color = Color.LightGreen;
            oSheet.get_Range("H23", "H23").Interior.Color = Color.LightCoral;

            //Kleur achtergrond kaders eindberekening
            oSheet.get_Range("S18", "S18").Interior.Color = Color.LightCoral;
            oSheet.get_Range("S19", "S19").Interior.Color = Color.LightGreen;
            oSheet.get_Range("S21", "S21").Interior.Color = Color.Gold;

            //Content messen
            oSheet.get_Range("A6", "A6").Value2 = lbl_Messen.Text;
            oSheet.get_Range("E6", "E6").Value2 = "-";
            oSheet.get_Range("F6", "F6").Value2 = Aantal_Messen;
            oSheet.get_Range("G6", "G6").Value2 = "5";
            oSheet.get_Range("H6", "H6").Value2 = WaardeMessen;

            //Content Oppervlakte
            oSheet.get_Range("A7", "A7").Value2 = lbl_Oppervlakte_Stickers.Text;
            oSheet.get_Range("F7", "F7").Value2 = Aantal;
            oSheet.get_Range("E7", "E7").Value2 = Oppervlakte;
            oSheet.get_Range("G7", "G7").Value2 = Prijs();
            oSheet.get_Range("H7", "H7").Value2 = WaardeOppervlakte * WaardeAantal * Prijs();

            //Middel 1
            oSheet.get_Range("A9", "A9").Value2 = lbl_Middel1.Text;
            oSheet.get_Range("F9", "F9").Value2 = txt_Middel1.Text;
            oSheet.get_Range("G9", "G9").Value2 = "13";
            oSheet.get_Range("H9", "H9").Value2 = Ver_WaardeMiddel1;

            oSheet.get_Range("E9", "E9").Value2 = "-";
            oSheet.get_Range("E9", "E9").HorizontalAlignment =
            Excel.XlHAlign.xlHAlignRight;

            //Middel 2
            oSheet.get_Range("A10", "A10").Value2 = lbl_Middel2.Text;
            oSheet.get_Range("F10", "F10").Value2 = txt_Middel2.Text;
            oSheet.get_Range("G10", "G10").Value2 = "23";
            oSheet.get_Range("H10", "H10").Value2 = Ver_WaardeMiddel2;

            oSheet.get_Range("E10", "E10").Value2 = "-";
            oSheet.get_Range("E10", "E10").HorizontalAlignment =
            Excel.XlHAlign.xlHAlignRight;

            //Middel 3
            oSheet.get_Range("A11", "A11").Value2 = lbl_Middel3.Text;
            oSheet.get_Range("F11", "F11").Value2 = txt_Middel3.Text;
            oSheet.get_Range("G11", "G11").Value2 = "29";
            oSheet.get_Range("H11", "H11").Value2 = Ver_WaardeMiddel3;

            oSheet.get_Range("E11", "E11").Value2 = "-";
            oSheet.get_Range("E11", "E11").HorizontalAlignment =
            Excel.XlHAlign.xlHAlignRight;

            //Verwijder kosten totaal
            oSheet.get_Range("H13", "H13").Value2 = WaardeMessen + (WaardeOppervlakte * WaardeAantal * Prijs()) + Ver_WaardeMiddel1 + Ver_WaardeMiddel2 + Ver_WaardeMiddel3;

            //Content Werkuren
            oSheet.get_Range("A16", "A16").Value2 = lbl_VerwachteTijdVerwijderen.Text;
            oSheet.get_Range("F16", "F16").Value2 = Ver_Werkuren;
            oSheet.get_Range("G16", "G16").Value2 = Ver_Prijs();
            oSheet.get_Range("H16", "H16").Value2 = Ver_WerkurenPrijs;
            #endregion

            #region Aanbrengen

            //Titels stickers aanbrengen
            oSheet.Cells[3, 12] = "Artikel soort";
            oSheet.Cells[3, 16] = "M²";
            oSheet.Cells[3, 17] = "Aantal";
            oSheet.Cells[3, 18] = "Prijs";
            oSheet.Cells[3, 19] = "Totaal";
            oSheet.Cells[5, 12] = "Kosten werkzaamheden";
            oSheet.Cells[13, 12] = "Aanbreng kosten totaal";
            oSheet.Cells[15, 12] = "Opbrengst werkzaamheden";

            oSheet.get_Range("L1", "L1").Value2 = "Stickers aanbrengen";

            //Content type materiaal
            oSheet.get_Range("L6", "L6").Value2 = lbl_Type_Materiaal.Text;
            oSheet.get_Range("P6", "P6").Value2 = txt_OppervlakteAanbrengen.Text;
            oSheet.get_Range("Q6", "Q6").Value2 = txt_AantalAanbrengen.Text;
            oSheet.get_Range("R6", "R6").Value2 = Aan_Materiaal();
            oSheet.get_Range("S6", "S6").Value2 = Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal();

            //Middel 1
            oSheet.get_Range("L9", "L9").Value2 = lbl_AanMiddel1.Text;
            oSheet.get_Range("Q9", "Q9").Value2 = txt_AanbrengenMiddel1.Text;
            oSheet.get_Range("R9", "R9").Value2 = "13";
            oSheet.get_Range("S9", "S9").Value2 = Aan_WaardeMiddel1;

            oSheet.get_Range("P9", "P9").Value2 = "-";
            oSheet.get_Range("P9", "P9").HorizontalAlignment =
            Excel.XlHAlign.xlHAlignRight;

            //Aanbreng kosten totaal
            oSheet.get_Range("S13", "S13").Value2 = Aan_WaardeMiddel1 + (Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal());

            //Content werkuren
            oSheet.get_Range("L16", "L16").Value2 = lbl_VerwachteTijdAanbrengen.Text;
            oSheet.get_Range("Q16", "Q16").Value2 = Aan_Werkuren;
            oSheet.get_Range("R16", "R16").Value2 = Aan_Prijs();
            oSheet.get_Range("S16", "S16").Value2 = Aan_WerkurenPrijs;


            #endregion

            #region Reiskosten

            //Reisvergoeding
            oSheet.get_Range("A20", "A20").Value2 = lbl_Reistijd.Text;
            oSheet.get_Range("E20", "E20").Value2 = WaardeReistijd;
            oSheet.get_Range("F20", "F20").Value2 = Werknemers();
            oSheet.get_Range("G20", "G20").Value2 = WaardeReis_Vergoeding;
            oSheet.get_Range("H20", "H20").Value2 = Werknemers() * WaardeReis_Vergoeding * WaardeReistijd;

            //Kosten reizen
            oSheet.get_Range("A23", "A23").Value2 = "Eigen vervoer";
            oSheet.get_Range("E23", "E23").Value2 = WaardeParkeerKosten;
            oSheet.get_Range("F23", "F23").Value2 = WaardePrijsBenzine;
            oSheet.get_Range("G23", "G23").Value2 = WaardeAfstand;
            oSheet.get_Range("H23", "H23").Value2 = WaardeParkeerKosten + (WaardePrijsBenzine * WaardeAfstand);

            #endregion

            #region Loonkosten

            oSheet.get_Range("W1", "W1").Value2 = "Loonkosten";

            oSheet.get_Range("W3", "W3").Value2 = "Naam werknemer";
            oSheet.get_Range("AA3", "AA3").Value2 = "LPU";
            oSheet.get_Range("AB3", "AB3").Value2 = "Uren";
            oSheet.get_Range("AC3", "AC3").Value2 = "Totaal";

            oSheet.get_Range("W5", "W5").Value2 = Naam1;
            oSheet.get_Range("AA5", "AA5").Value2 = WaardeLPU1;
            oSheet.get_Range("AB5", "AB5").Value2 = WaardeTotaleWerkuren1;
            oSheet.get_Range("AC5", "AC5").Value2 = WaardeLPU1 * WaardeTotaleWerkuren1;

            oSheet.get_Range("W6", "W6").Value2 = Naam2;
            oSheet.get_Range("AA6", "AA6").Value2 = WaardeLPU2;
            oSheet.get_Range("AB6", "AB6").Value2 = WaardeTotaleWerkuren2;
            oSheet.get_Range("AC6", "AC6").Value2 = WaardeLPU2 * WaardeTotaleWerkuren2;

            oSheet.get_Range("W7", "W7").Value2 = Naam3;
            oSheet.get_Range("AA7", "AA7").Value2 = WaardeLPU3;
            oSheet.get_Range("AB7", "AB7").Value2 = WaardeTotaleWerkuren3;
            oSheet.get_Range("AC7", "AC7").Value2 = WaardeLPU3 * WaardeTotaleWerkuren3;

            oSheet.get_Range("W8", "W8").Value2 = Naam4;
            oSheet.get_Range("AA8", "AA8").Value2 = WaardeLPU4;
            oSheet.get_Range("AB8", "AB8").Value2 = WaardeTotaleWerkuren4;
            oSheet.get_Range("AC8", "AC8").Value2 = WaardeLPU4 * WaardeTotaleWerkuren4;
            

            #endregion

            #region Eindberekening

            oSheet.get_Range("S18", "S18").Value2 = (WaardeMessen + (WaardeOppervlakte * WaardeAantal * Prijs()) + Ver_WaardeMiddel1 + Ver_WaardeMiddel2 + Ver_WaardeMiddel3)
                                                  + (Aan_WaardeMiddel1 + (Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal()))
                                                  + (WaardeParkeerKosten + (WaardePrijsBenzine * WaardeAfstand));


            oSheet.get_Range("S19", "S19").Value2 = (Ver_WerkurenPrijs)
                                                  + (Aan_WerkurenPrijs)
                                                  + (Werknemers() * WaardeReis_Vergoeding * WaardeReistijd);

            oSheet.get_Range("S21", "S21").Value2 = ((Ver_WerkurenPrijs)
                                                  + (Aan_WerkurenPrijs)
                                                  + (Werknemers() * WaardeReis_Vergoeding * WaardeReistijd)) 
                                                  +
                                                  ((WaardeMessen + (WaardeOppervlakte * WaardeAantal * Prijs()) + Ver_WaardeMiddel1 + Ver_WaardeMiddel2 + Ver_WaardeMiddel3)
                                                  + (Aan_WaardeMiddel1 + (Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal()))
                                                  + (WaardeParkeerKosten + (WaardePrijsBenzine * WaardeAfstand)));
            #endregion

            #region Extra Info

            oSheet.get_Range("A28", "A28").Value2 = "Type klus:";
            oSheet.get_Range("E28", "E28").Value2 = Type_Klus;
            oSheet.get_Range("A29", "A29").Value2 = "Order nummer:";
            oSheet.get_Range("E29", "E29").Value2 = WaardeOrderNummer;

            oSheet.get_Range("A31", "A31").Value2 = "Datum:";
            oSheet.get_Range("E31", "E31").Value2 = Datum;
            oSheet.get_Range("A32", "A32").Value2 = "Locatie:";
            oSheet.get_Range("E32", "E32").Value2 = Locatie;
            oSheet.get_Range("A33", "A33").Value2 = "Temperatuur:";
            oSheet.get_Range("E33", "E33").Value2 = WaardeTemperatuur;

            oSheet.get_Range("A35", "A35").Value2 = "Voorbereidingstijd:";
            oSheet.get_Range("E35", "E35").Value2 = WaardeVoorbereidingstijd;
            oSheet.get_Range("A36", "A36").Value2 = "Totale duur werkzaamheden:";
            oSheet.get_Range("E36", "E36").Value2 = Waarde_Ver_Werkuren + Aan_Waarde_Werkuren;
            oSheet.get_Range("A37", "A37").Value2 = "Geschat aantal werkdagen:";
            oSheet.get_Range("E37", "E37").Value2 = (Waarde_Ver_Werkuren + Aan_Waarde_Werkuren)/8;

            oSheet.get_Range("A40", "A40").Value2 = "Opmerkingen:";
            oSheet.get_Range("A40", "A40").Value2 = Opmerkingen;

            oSheet.get_Range("E29", "E29").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            oSheet.get_Range("E33", "E33").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            //oRng = oSheet.get_Range("A28", "S35");
            //oRng.EntireColumn.AutoFit();


            #endregion
        }

        private void txt_Oppervlakte_TextChanged(object sender, EventArgs e)
        {
            InvoerCheck(txt_Oppervlakte, false);
        }
        #endregion


        /// <summary>
        /// Functie om de invoer van de textboxen te controleren en te verkleuren.
        /// </summary>
        /// <param name="textBox">Invoer eigenschappen opdracht</param>
        public void InvoerCheck(System.Windows.Forms.TextBox textBox, bool isletter)
        {
            if (isletter)
            {
                if (!(string.IsNullOrEmpty(textBox.Text)) && textBox.Text.All(Char.IsLetter))
                {
                    textBox.BackColor = Color.LightGreen;
                }
                else if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.BackColor = Color.White;
                }
                else
                {
                    textBox.BackColor = Color.LightCoral;
                }
            }
            else
            {
                if (!(string.IsNullOrEmpty(textBox.Text)) && textBox.Text.All(Char.IsDigit) || textBox.Text.Contains(",") && !textBox.Text.Any(Char.IsLetter))
                {
                    textBox.BackColor = Color.LightGreen;
                }
                else if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.BackColor = Color.White;
                }
                else
                {
                    textBox.BackColor = Color.LightCoral;
                }
            }

        }

        /// <summary>
        /// Functie om de invoer van de comboboxen te controleren en te verkleuren.
        /// </summary>
        /// <param name="textBox">Invoer eigenschappen opdracht</param>
        public void InvoerCheckCmb(System.Windows.Forms.ComboBox textBox, bool isletter)
        {
            if (isletter)
            {
                if (!(string.IsNullOrEmpty(textBox.Text)) && textBox.Text.All(Char.IsLetter))
                {
                    textBox.BackColor = Color.LightGreen;
                }
                else if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.BackColor = Color.White;
                }
                else
                {
                    textBox.BackColor = Color.LightGreen;
                }
            }
            else
            {
                if (!(string.IsNullOrEmpty(textBox.Text)) && textBox.Text.All(Char.IsDigit) || textBox.Text.Contains(",") && !textBox.Text.Any(Char.IsLetter))
                {
                    textBox.BackColor = Color.LightGreen;
                }
                else if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.BackColor = Color.White;
                }
                else
                {
                    textBox.BackColor = Color.LightGreen;
                }
            }

        }

        #region ComboBox eenheden
        public double Prijs()
        {
            switch (cmb_Prijs_Per_m2.Text)
            {
                case "5":
                    return 5;
                break;

                case "10":
                    return 10;
                    break;

                case "20":
                    return 20;
                    break;

                default: return 0;
            }
            
        }
        public double Ver_Prijs()
        {
            switch (cmb_PPU_Verwijderen.Text)
            {
                case "12,50":
                    return 12.50;
                break;

                case "15":
                    return 15;
                    break;

                case "23,50":
                    return 23.50;
                    break;

                case "35":
                    return 35;
                    break;

                default: return 0;
            }
        }

        public double Aan_Prijs()
        {
            switch (cmb_PPU_Aanbrengen.Text)
            {
                case "12,50":
                    return 12.50;
                    break;

                case "15":
                    return 15;
                    break;

                case "23,50":
                    return 23.50;
                    break;

                case "35":
                    return 35;
                    break;

                default: return 0;
            }
        }

        public double Aan_Materiaal()
        {
            switch (cmb_Type_Materiaal.Text)
            {
                case "Staal (12,75)":
                    return 12.75;
                    break;

                case "Plastic (9,25)":
                    return 9.25;
                    break;

                case "Hout (6,90)":
                    return 6.90;
                    break;

                default: return 0;
            }
        }

        public double Werknemers()
        {
            switch (cmb_AantalWerknemers.Text)
            {
                case "1":
                    return 1;
                    break;

                case "2":
                    return 2;
                    break;

                case "3":
                    return 3;
                    break;

                case "5":
                    return 5;

                case "10":
                    return 10;

                default: return 0;
            }
        }
        #endregion

        //loonberekening test
        public void Loonberekening()
        {
            Naam1 = txt_Naam1.Text;
            Naam2 = txt_Naam2.Text;
            Naam3 = txt_Naam3.Text;
            Naam4 = txt_Naam4.Text;

            LPU1 = txt_LPU1.Text;
            LPU2 = txt_LPU2.Text;
            LPU3 = txt_LPU3.Text;
            LPU4 = txt_LPU4.Text;

            TotaleWerkuren1 = txt_TotaleWerkuren1.Text;
            TotaleWerkuren2 = txt_TotaleWerkuren2.Text;
            TotaleWerkuren3 = txt_TotaleWerkuren3.Text;
            TotaleWerkuren4 = txt_TotaleWerkuren4.Text;

            if (LPU1.All(Char.IsDigit) && TotaleWerkuren1.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_LPU1.Text))
                    && !(string.IsNullOrEmpty(txt_TotaleWerkuren1.Text))
                    

                    || LPU1.Contains(",") && !LPU1.Any(Char.IsLetter)
                    || TotaleWerkuren1.Contains(",") && !(TotaleWerkuren1.Any(Char.IsLetter))

                    )

            {
                WaardeLPU1 = Convert.ToDouble(LPU1);
                WaardeTotaleWerkuren1 = Convert.ToDouble(TotaleWerkuren1);
            }

            if (LPU2.All(Char.IsDigit) && TotaleWerkuren2.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_LPU2.Text))
                    && !(string.IsNullOrEmpty(txt_TotaleWerkuren2.Text))


                    || LPU2.Contains(",") && !LPU2.Any(Char.IsLetter)
                    || TotaleWerkuren2.Contains(",") && !(TotaleWerkuren2.Any(Char.IsLetter))

                    )

            {
                WaardeLPU2 = Convert.ToDouble(LPU2);
                WaardeTotaleWerkuren2 = Convert.ToDouble(TotaleWerkuren2);
            }

            if (LPU3.All(Char.IsDigit) && TotaleWerkuren3.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_LPU3.Text))
                    && !(string.IsNullOrEmpty(txt_TotaleWerkuren3.Text))


                    || LPU3.Contains(",") && !LPU3.Any(Char.IsLetter)
                    || TotaleWerkuren3.Contains(",") && !(TotaleWerkuren3.Any(Char.IsLetter))

                    )

            {
                WaardeLPU3 = Convert.ToDouble(LPU3);
                WaardeTotaleWerkuren3 = Convert.ToDouble(TotaleWerkuren3);
            }

            if (LPU4.All(Char.IsDigit) && TotaleWerkuren4.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_LPU4.Text))
                    && !(string.IsNullOrEmpty(txt_TotaleWerkuren4.Text))


                    || LPU4.Contains(",") && !LPU4.Any(Char.IsLetter)
                    || TotaleWerkuren4.Contains(",") && !(TotaleWerkuren4.Any(Char.IsLetter))

                    )

            {
                WaardeLPU4 = Convert.ToDouble(LPU4);
                WaardeTotaleWerkuren4 = Convert.ToDouble(TotaleWerkuren4);
            }
        }

        public void EigenschappenOpdrachtberekening()
        {
            Type_Klus = txt_Type_Klus.Text;
            Locatie = txt_locatie.Text;
            Datum = dtp_Datum.Text;

            OrderNummer = txt_OrderNummer.Text;
            Temperatuur = txt_Temperatuur.Text;
            Voorbereidingstijd = txt_Voorbereidingstijd.Text;


            if (OrderNummer.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_OrderNummer.Text))

                    || OrderNummer.Contains(",") && !OrderNummer.Any(Char.IsLetter)

                    )

            {
                WaardeOrderNummer = Convert.ToDouble(OrderNummer);
            }

            if (Temperatuur.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_Temperatuur.Text))

                    || Temperatuur.Contains(",") && !Temperatuur.Any(Char.IsLetter)

                    )

            {
                WaardeTemperatuur = Convert.ToDouble(Temperatuur);
            }

            if (Voorbereidingstijd.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_Voorbereidingstijd.Text))

                    || Voorbereidingstijd.Contains(",") && !Voorbereidingstijd.Any(Char.IsLetter)

                    )

            {
                WaardeVoorbereidingstijd = Convert.ToDouble(Voorbereidingstijd);
            }
        }

        public void TotaleKosten()
        {
            EindberekeningKosten = (WaardeMessen + (WaardeOppervlakte * WaardeAantal * Prijs()) + Ver_WaardeMiddel1 + Ver_WaardeMiddel2 + Ver_WaardeMiddel3)
                                                  + (Aan_WaardeMiddel1 + (Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal()))
                                                  + (WaardeParkeerKosten + (WaardePrijsBenzine * WaardeAfstand));

            EindberekeningVerdiensten =             (Ver_WerkurenPrijs)
                                                  + (Aan_WerkurenPrijs)
                                                  + (Werknemers() * WaardeReis_Vergoeding * WaardeReistijd);

            EindberekeningWinst =                   ((Ver_WerkurenPrijs)
                                                  + (Aan_WerkurenPrijs)
                                                  + (Werknemers() * WaardeReis_Vergoeding * WaardeReistijd))
                                                  +
                                                  ((WaardeMessen + (WaardeOppervlakte * WaardeAantal * Prijs()) + Ver_WaardeMiddel1 + Ver_WaardeMiddel2 + Ver_WaardeMiddel3)
                                                  + (Aan_WaardeMiddel1 + (Aan_WaardeOppervlakte * Aan_WaardeAantal * Aan_Materiaal()))
                                                  + (WaardeParkeerKosten + (WaardePrijsBenzine * WaardeAfstand)));
        }

        public void OverigeKosten()
        {
            RollenPoetspapier = txt_RolPapier.Text;
            ExtraBenodigdheden = txt_OverigeKosten.Text;

            if (RollenPoetspapier.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_RolPapier.Text))

                    || RollenPoetspapier.Contains(",") && !RollenPoetspapier.Any(Char.IsLetter)

                    )

            {
                WaardeRollenPoetspapier = Convert.ToDouble(RollenPoetspapier);
            }

            if (ExtraBenodigdheden.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_OverigeKosten.Text))

                    || ExtraBenodigdheden.Contains(",") && !ExtraBenodigdheden.Any(Char.IsLetter)

                    )

            {
                WaardeExtraBenodigdheden = Convert.ToDouble(ExtraBenodigdheden);
            }
        }

        private void btn_Berekenen_Click(object sender, EventArgs e)
        {
            Loonberekening();
            EigenschappenOpdrachtberekening();
            OverigeKosten();
            
            //Verwijderen invoer naar string
            Oppervlakte = txt_Oppervlakte.Text;
            Aantal = txt_Aantal.Text;
            Type_Klus = txt_Type_Klus.Text;
            Aantal_Messen = txt_Messen.Text;
            Ver_Werkuren = txt_WerkurenVerwijderen.Text;

            Ver_Middel1 = txt_Middel1.Text;
            Ver_Middel2 = txt_Middel2.Text;
            Ver_Middel3 = txt_Middel3.Text;

            //Aanbrengen invoer naar string
            Aan_Oppervlakte = txt_OppervlakteAanbrengen.Text;
            Aan_Aantal = txt_AantalAanbrengen.Text;
            Aan_Middel1 = txt_AanbrengenMiddel1.Text;
            Aan_Werkuren = txt_WerkurenAanbrengen.Text;

            //Reisvergoeding invoer naar string
            Reistijd = txt_Reistijd.Text;
            Reis_Vergoeding = txt_ReisVergoeding.Text;

            //Reiskosten invoer naar string
            PrijsBenzine = txt_PrijsBenzine.Text;
            ParkeerKosten = txt_ParkeerKosten.Text;
            Afstand = txt_Afstand.Text;

            Datum = dtp_Datum.Text;
            Opmerkingen = txt_Opmerkingen.Text;

            void Error_msg()
            {
                MessageBox.Show("Voer een getal in i.p.v. letters en zorg dat alle velden zijn ingevuld");
            }

            //Controle invoer stickers verwijderen
            if (
                Aantal.All(Char.IsDigit) && Oppervlakte.All(Char.IsDigit) && Aantal_Messen.All(Char.IsDigit) && Ver_Werkuren.All(Char.IsDigit) && Ver_Middel1.All(Char.IsDigit) && Ver_Middel2.All(Char.IsDigit) && Ver_Middel3.All(Char.IsDigit) && Reistijd.All(Char.IsDigit) && Reis_Vergoeding.All(Char.IsDigit)

                && !(string.IsNullOrEmpty(txt_Messen.Text))
                && !(string.IsNullOrEmpty(txt_Oppervlakte.Text))
                && !(string.IsNullOrEmpty(txt_Aantal.Text))
                && !(string.IsNullOrEmpty(txt_Type_Klus.Text))
                && !(string.IsNullOrEmpty(cmb_Prijs_Per_m2.Text))
                && !(string.IsNullOrEmpty(txt_WerkurenVerwijderen.Text))
                && !(string.IsNullOrEmpty(cmb_PPU_Verwijderen.Text))
                && !(string.IsNullOrEmpty(txt_Middel1.Text))
                && !(string.IsNullOrEmpty(txt_Middel2.Text))
                && !(string.IsNullOrEmpty(txt_Middel3.Text))
                && !(string.IsNullOrEmpty(txt_Reistijd.Text))
                && !(string.IsNullOrEmpty(txt_ReisVergoeding.Text))

                || Aantal.Contains(",") && !Aantal.Any(Char.IsLetter)
                || Oppervlakte.Contains(",") && !(Oppervlakte.Any(Char.IsLetter))
                || Aantal_Messen.Contains(",") && !(Aantal_Messen.Any(Char.IsLetter))
                || Ver_Werkuren.Contains(",") && !(Ver_Werkuren.Any(Char.IsLetter))
                || Ver_Middel1.Contains(",") && !(Ver_Middel1.Any(Char.IsLetter))
                || Ver_Middel2.Contains(",") && !(Ver_Middel2.Any(Char.IsLetter))
                || Ver_Middel3.Contains(",") && !(Ver_Middel3.Any(Char.IsLetter))
                || Reistijd.Contains(",") && !(Reistijd.Any(Char.IsLetter))
                || Reis_Vergoeding.Contains(",") && !(Reis_Vergoeding.Any(Char.IsLetter))
                )
            {
                WaardeAantal = Convert.ToDouble(Aantal);
                WaardeOppervlakte = Convert.ToDouble(Oppervlakte);
                WaardeMessen = Convert.ToDouble(Aantal_Messen) * 5;
                Waarde_Ver_Werkuren = Convert.ToDouble(Ver_Werkuren);

                Ver_WaardeMiddel1 = Convert.ToDouble(Ver_Middel1) * 13;
                Ver_WaardeMiddel2 = Convert.ToDouble(Ver_Middel2) * 23;
                Ver_WaardeMiddel3 = Convert.ToDouble(Ver_Middel3) * 29;

                WaardeReistijd = Convert.ToDouble(Reistijd);
                WaardeReis_Vergoeding = Convert.ToDouble(Reis_Vergoeding);

                //Controle of stickers aanbrengen ook is ingevoerd
                if (Aan_Oppervlakte.All(Char.IsDigit) && Aan_Aantal.All(Char.IsDigit) && Aan_Middel1.All(Char.IsDigit) && Aan_Werkuren.All(Char.IsDigit) && Reistijd.All(Char.IsDigit) && Reis_Vergoeding.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_OppervlakteAanbrengen.Text))
                    && !(string.IsNullOrEmpty(txt_AantalAanbrengen.Text))
                    && !(string.IsNullOrEmpty(txt_AanbrengenMiddel1.Text))
                    && !(string.IsNullOrEmpty(txt_WerkurenAanbrengen.Text))
                    && !(string.IsNullOrEmpty(txt_Reistijd.Text))
                    && !(string.IsNullOrEmpty(txt_ReisVergoeding.Text))

                    || Aan_Oppervlakte.Contains(",") && !Aan_Oppervlakte.Any(Char.IsLetter)
                    || Aan_Aantal.Contains(",") && !Aan_Aantal.Any(Char.IsLetter)
                    || Aan_Middel1.Contains(",") && !Aan_Middel1.Any(Char.IsLetter)
                    || Aan_Werkuren.Contains(",") && !Aan_Werkuren.Any(Char.IsLetter)
                    || Reistijd.Contains(",") && !(Reistijd.Any(Char.IsLetter))
                    || Reis_Vergoeding.Contains(",") && !(Reis_Vergoeding.Any(Char.IsLetter))
                    )

                {
                    Aan_WaardeOppervlakte = Convert.ToDouble(Aan_Oppervlakte);
                    Aan_WaardeAantal = Convert.ToDouble(Aan_Aantal);
                    Aan_Waarde_Werkuren = Convert.ToDouble(Aan_Werkuren);

                    Aan_WaardeMiddel1 = Convert.ToDouble(Aan_Middel1) * 13;

                    WaardeReistijd = Convert.ToDouble(Reistijd);
                    WaardeReis_Vergoeding = Convert.ToDouble(Reis_Vergoeding);

                    //Controle invoer soort vervoer
                    if (PrijsBenzine.All(Char.IsDigit) && ParkeerKosten.All(Char.IsDigit) && Afstand.All(Char.IsDigit)

                        && !(string.IsNullOrEmpty(txt_PrijsBenzine.Text))
                        && !(string.IsNullOrEmpty(txt_ParkeerKosten.Text))
                        && !(string.IsNullOrEmpty(txt_Afstand.Text))

                        || PrijsBenzine.Contains(",") && !PrijsBenzine.Any(Char.IsLetter)
                        || ParkeerKosten.Contains(",") && !ParkeerKosten.Any(Char.IsLetter)
                        || Afstand.Contains(",") && !Afstand.Any(Char.IsLetter)
                        )
                    {
                        WaardePrijsBenzine = Convert.ToDouble(PrijsBenzine);
                        WaardeParkeerKosten = Convert.ToDouble(ParkeerKosten);
                        WaardeAfstand = Convert.ToDouble(Afstand);
                    }
                }
                //Controle invoer soort vervoer
                if (PrijsBenzine.All(Char.IsDigit) && ParkeerKosten.All(Char.IsDigit) && Afstand.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_PrijsBenzine.Text))
                    && !(string.IsNullOrEmpty(txt_ParkeerKosten.Text))
                    && !(string.IsNullOrEmpty(txt_Afstand.Text))

                    || PrijsBenzine.Contains(",") && !PrijsBenzine.Any(Char.IsLetter)
                    || ParkeerKosten.Contains(",") && !ParkeerKosten.Any(Char.IsLetter)
                    || Afstand.Contains(",") && !Afstand.Any(Char.IsLetter)
                    )
                {
                    WaardePrijsBenzine = Convert.ToDouble(PrijsBenzine);
                    WaardeParkeerKosten = Convert.ToDouble(ParkeerKosten);
                    WaardeAfstand = Convert.ToDouble(Afstand);
                }
            }
            
            //Controle invoer stickers aanbrengen
            else if (Aan_Oppervlakte.All(Char.IsDigit) && Aan_Aantal.All(Char.IsDigit) && Aan_Middel1.All(Char.IsDigit) && Aan_Werkuren.All(Char.IsDigit) && Reistijd.All(Char.IsDigit) && Reis_Vergoeding.All(Char.IsDigit)

                && !(string.IsNullOrEmpty(txt_OppervlakteAanbrengen.Text))
                && !(string.IsNullOrEmpty(txt_AantalAanbrengen.Text))
                && !(string.IsNullOrEmpty(txt_AanbrengenMiddel1.Text))
                && !(string.IsNullOrEmpty(txt_WerkurenAanbrengen.Text))
                && !(string.IsNullOrEmpty(txt_Reistijd.Text))
                && !(string.IsNullOrEmpty(txt_ReisVergoeding.Text))

                || Aan_Oppervlakte.Contains(",") && !Aan_Oppervlakte.Any(Char.IsLetter)
                || Aan_Aantal.Contains(",") && !Aan_Aantal.Any(Char.IsLetter)
                || Aan_Middel1.Contains(",") && !Aan_Middel1.Any(Char.IsLetter)
                || Aan_Werkuren.Contains(",") && !Aan_Werkuren.Any(Char.IsLetter)
                || Reistijd.Contains(",") && !(Reistijd.Any(Char.IsLetter))
                || Reis_Vergoeding.Contains(",") && !(Reis_Vergoeding.Any(Char.IsLetter))
                )
            {
                Aan_WaardeOppervlakte = Convert.ToDouble(Aan_Oppervlakte);
                Aan_WaardeAantal = Convert.ToDouble(Aan_Aantal);
                Aan_Waarde_Werkuren = Convert.ToDouble(Aan_Werkuren);

                Aan_WaardeMiddel1 = Convert.ToDouble(Aan_Middel1) * 13;

                WaardeReistijd = Convert.ToDouble(Reistijd);
                WaardeReis_Vergoeding = Convert.ToDouble(Reis_Vergoeding);
                
                //Controle invoer soort vervoer
                if (PrijsBenzine.All(Char.IsDigit) && ParkeerKosten.All(Char.IsDigit) && Afstand.All(Char.IsDigit)

                    && !(string.IsNullOrEmpty(txt_PrijsBenzine.Text))
                    && !(string.IsNullOrEmpty(txt_ParkeerKosten.Text))
                    && !(string.IsNullOrEmpty(txt_Afstand.Text))

                    || PrijsBenzine.Contains(",") && !PrijsBenzine.Any(Char.IsLetter)
                    || ParkeerKosten.Contains(",") && !ParkeerKosten.Any(Char.IsLetter)
                    || Afstand.Contains(",") && !Afstand.Any(Char.IsLetter)
                    )
                {
                    WaardePrijsBenzine = Convert.ToDouble(PrijsBenzine);
                    WaardeParkeerKosten = Convert.ToDouble(ParkeerKosten);
                    WaardeAfstand = Convert.ToDouble(Afstand);
                }

            }
            else
            {
                Error_msg();
            }
            //Character_Check(Oppervlakte, WaardeOppervlakte);
            //Character_Check(WaardeAantal = Convert.ToDouble(Aantal)); 

            double WaardeTotaal = WaardeOppervlakte * WaardeAantal * Prijs() + WaardeMessen;
            Ver_WerkurenPrijs = Ver_Prijs() * Waarde_Ver_Werkuren;
            Aan_WerkurenPrijs = Aan_Prijs() * Aan_Waarde_Werkuren;

            WaardeReiskosten = Werknemers() * WaardeReistijd * WaardeReis_Vergoeding;


            lbl_TotaalVerwijderen.Text = "Totale kosten verwijderen: €" + WaardeTotaal;
            lbl_ResultaatVerOppervlakte.Text = "Totale oppervlakte (in m²): " + (WaardeAantal * WaardeOppervlakte);

            //lbl_TotaalAanbrengen.Text = "Totale kosten Aanbreng: €" + 
            lbl_ResultaatAanOppervlakte.Text = "Totale oppervlakte (in m²): " + (Aan_WaardeAantal * Aan_WaardeOppervlakte);

            lbl_ResultaatLoon.Text = "Totale loon kosten: " + ((WaardeLPU1 * WaardeTotaleWerkuren1) + (WaardeLPU2 * WaardeTotaleWerkuren2) + (WaardeLPU3 * WaardeTotaleWerkuren3) + (WaardeLPU4 * WaardeTotaleWerkuren4));
            lbl_ResultaatWerkuren.Text = "Totale werkuren: " + (Waarde_Ver_Werkuren + Aan_Waarde_Werkuren);

            lbl_Resultaat_Type_Klus.Text = "Type klus: " + Type_Klus;
            lbl_ResultaatDatum.Text = "Datum: " + Datum;
            lbl_ResultaatLocatie.Text = "Locatie: " + Locatie;
            lbl_ResultaatDuurOpdracht.Text = "Totale duur opdracht: " + (Aan_Waarde_Werkuren + Waarde_Ver_Werkuren);
            lbl_ResultaatWerkdagen.Text = "Geschat aantal werkdagen: " + ((Waarde_Ver_Werkuren + Aan_Waarde_Werkuren) / 8);

            //Activatie labels na berekening
            lbl_TotaalVerwijderen.Visible = true;
            lbl_ResultaatVerOppervlakte.Visible = true;

            lbl_TotaalAanbrengen.Visible = true;
            lbl_ResultaatAanOppervlakte.Visible = true;

            lbl_ResultaatLoon.Visible = true;
            lbl_ResultaatWerkuren.Visible = true;

            lbl_Resultaat_Type_Klus.Visible = true;
            lbl_ResultaatDatum.Visible = true;
            lbl_ResultaatLocatie.Visible = true;
            lbl_ResultaatDuurOpdracht.Visible = true;
            lbl_ResultaatWerkdagen.Visible = true;

            //Berekening eindkosten
            TotaleKosten();

            lbl_ResultaatKosten.Text = "Totale kosten: " + EindberekeningKosten;


            lbl_ResultaatVerdiensten.Text = "Totale verdiensten: " + EindberekeningVerdiensten;

            lbl_ResultaatWinst.Text = "Totale winst: " + EindberekeningWinst;

            //Activatie labels eindwaarden
            lbl_ResultaatVerdiensten.Visible = true;
            lbl_ResultaatKosten.Visible = true;
            lbl_ResultaatWinst.Visible = true;
            pnl_Resultaat.Visible = true;
        }
    }
}
