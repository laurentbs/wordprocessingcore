using System;
using System.IO;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
using Telerik.Windows.Documents.Flow.Model.Styles;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Media;
using System.Linq;
using System.Collections.Generic;

namespace WordsProcessingCore
{
    class Program
    {
        static void Main(string[] args)
        {
            // Affiche hello world sur la console
            Console.WriteLine("Hello World!");

            // Crée un nouveau document "radflow" représenté par la variable document1 et y insère le text "Hello world!"
            RadFlowDocument document1 = new();
            RadFlowDocumentEditor editor = new(document1);
            editor.InsertText("Hello world!");

            // Insère une table vide dans le document
            Table table0 = editor.InsertTable(5, 3);

            // Insère une table, puis génère du contenu de cellules
            Table table = document1.Sections.AddSection().Blocks.AddTable();
            document1.StyleRepository.AddBuiltInStyle(BuiltInStyleNames.TableGridStyleId);
            table.StyleId = BuiltInStyleNames.TableGridStyleId;

            ThemableColor cellBackground = new ThemableColor(Colors.Gray);

            for (int i = 0; i < 5; i++)
            {
                TableRow row = table.Rows.AddTableRow();
                for (int j = 0; j < 10; j++)
                {
                    TableCell cell = row.Cells.AddTableCell();
                    cell.Blocks.AddParagraph().Inlines.AddRun(string.Format("Cell {0}, {1}", i, j));
                    cell.Shading.BackgroundColor = cellBackground;
                    cell.PreferredWidth = new TableWidthUnit(50);
                }
            }

            // Sauvegarde le document au format docx
            using (Stream output = new FileStream("output.docx", FileMode.OpenOrCreate))
            {
                DocxFormatProvider provider1 = new();
                provider1.Export(document1, output);
            }

            // Ouvre le document en lecture et importe le comme un objet radflow
            DocxFormatProvider provider2 = new();
            using Stream input = File.OpenRead("input.docx");
            RadFlowDocument document2 = provider2.Import(input);

            // Liste toutes les tables dans le document2
            List<Table> tables = document2.EnumerateChildrenOfType<Table>().ToList();
            // Boucle dans chacune des tables de la liste
            foreach (Table myTable in tables)
            {
                // Boucle dans chacune des lignes de la table
                foreach (TableRow myRow in myTable.Rows)
                {
                    // Boucle dans chacune des cellules de la ligne
                    foreach (TableCell myCell in myRow.Cells)
                    {
                        // Liste tous les "runs" de texte dans la cellule
                        List<Run> runs = myCell.EnumerateChildrenOfType<Run>().ToList();
                        // Boucle dans chaque run de texte
                        foreach (Run myRun in runs)
                        {
                            // Affiche le contenu de chaque run de texte (=contenu texte de la cellule)
                            Console.WriteLine(myRun.Text);
                        }
                    }
                }
            }
        }
    }
}
