using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace LerPlanilhaExcel
{
    class Program
    {

        static List<string> numeroContrato = new List<string>();

        static void Main(string[] args)
        {
            readPlanilha("janeiro_2020");
            readPlanilha("fevereiro_2020");
            readPlanilha("marco_2020");
            readPlanilha("abril_2020");
            readPlanilha("maio_2020");
            readPlanilha("junho_2020");
            readPlanilha("julho_2020");
            readPlanilha("agosto_2020");
            readPlanilha("setembro_2020");
            readPlanilha("outubro_2020");
            readPlanilha("novembro_2020");
            readPlanilha("dezembro_2020");
            readPlanilha("janeiro_2021");
            readPlanilha("fevereiro_2021");
            readPlanilha("marco_2021");
            readPlanilha("abril_2021");
            readPlanilha("maio_2021");
            readPlanilha("junho_2021");
            readPlanilha("julho_2021");
            readPlanilha("agosto_2021");
            readPlanilha("setembro_2021");
            readPlanilha("outubro_2021");
        }

        static void readPlanilha(string nomePlanilha){
            var xls = new XLWorkbook(@"C:\Users\Qualidade\Documents\TarefaOPNovo\204\contratos_ponta_pora_atualizacao_nome_contrato_2021_10_2021.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == nomePlanilha);
            var totalLinhas = planilha.Rows().Count();
            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                /*
                var codigo = int.Parse(planilha.Cell($"A{l}").Value.ToString());
                var descricao  = planilha.Cell($"B{l}").Value.ToString();
                var preco = decimal.Parse(planilha.Cell($"C{l}").Value.ToString());
                Console.WriteLine($"{codigo} - {descricao} - {preco}");
                */
                var A = planilha.Cell($"A{l}").Value.ToString();
                var B  = planilha.Cell($"B{l}").Value.ToString();
                var C = planilha.Cell($"C{l}").Value.ToString();

                bool flagExist = false;
                foreach(var num in numeroContrato){
                    if(num==C){
                        flagExist = true;
                    }
                }
                if(!flagExist){
                    numeroContrato.Add(C);
                    writeFile(A, B, C);
                }
            }
        }

        static void writeFile(string A, string B, string C){
            try
            {
                //Pass the filepath and filename to the StreamWriter Constructor
                //StreamWriter sw = new StreamWriter(@"C:\Users\Qualidade\Documents\TarefaOPNovo\204\Test.txt");
                //Write a line of text
                //sw.WriteLine(A + ", " + B + ", " + C);

                string path = @"C:\Users\Qualidade\Documents\TarefaOPNovo\204\GeradoAutomaticamente.txt";

                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("\nUPDATE public." + "\"" + "ContratosFornecedores" + "\"" +
                                 "\n    SET " + "\"" + "NomeContrato" + "\"" + $"='{B}'" +
                                 "\n    WHERE " + "\"" + "NumeroContrato" + "\"" + $"='{C}';");

                    sw.Close();
                }	

                //Write a second line of text
                //sw.WriteLine("From the StreamWriter class");
                //Close the file
                //sw.Close();
            }
            catch(Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }
        }
    }
}