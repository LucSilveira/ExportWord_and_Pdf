using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;

namespace Imprimindo_dados_usuario
{
    class Program
    {
        static void Main(string[] args)
        {

            string nomes;
            string enderecos;
            double compra;
            string data;

                Console.WriteLine("Informe seu nome por favor:");
                nomes = Console.ReadLine();

                Console.WriteLine("Informe seu endereço:");
                enderecos = Console.ReadLine();

                Console.WriteLine("Informe o valor da sua compra");
                compra = double.Parse(Console.ReadLine());

                Console.WriteLine("Informe a data da sua compra");
                data = DateTime.Parse(Console.ReadLine()).ToString("dd-MM-yyyy");


            Document documentImpressao = new Document();

            Section secaoCapa = documentImpressao.AddSection();
            Paragraph notaFiscal = secaoCapa.AddParagraph();

            notaFiscal.AppendText("Nota Fiscal\n");

            
            CharacterFormat format = new CharacterFormat(documentImpressao);
            format.Bold = true;

                Paragraph nome = secaoCapa.AddParagraph();
                nome.AppendText("Nome: ") .ApplyCharacterFormat(format);

                // Paragraph nomeUsuario = secaoCapa.AddParagraph();
                nome.AppendText($"{nomes}\n");

                Paragraph endereco = secaoCapa.AddParagraph();
                endereco.AppendText("Endereço: ").ApplyCharacterFormat(format);

                // Paragraph nomeEndereco = secaoCapa.AddParagraph();
                endereco.AppendText($"{enderecos}\n");

                Paragraph valor = secaoCapa.AddParagraph();
                valor.AppendText("Valor: ").ApplyCharacterFormat(format);

                // Paragraph valorCompra = secaoCapa.AddParagraph();
                valor.AppendText($"{compra}\n");

                Paragraph dataDaCompra = secaoCapa.AddParagraph();
                dataDaCompra.AppendText("Data: ").ApplyCharacterFormat(format);

                // Paragraph dataCompra = secaoCapa.AddParagraph();
                dataDaCompra.AppendText($"{data}\n");


            // ParagraphStyle estiloParafrafo = new ParagraphStyle(documentImpressao);
            // estiloParafrafo.Name = "Cor do titulo";
            // estiloParafrafo.CharacterFormat.TextColor = Color.DarkBlue;
            // estiloParafrafo.CharacterFormat.Bold = true;
            // documentImpressao.Styles.Add(estiloParafrafo);
            // notaFiscal.ApplyStyle(estiloParafrafo.Name);
            // nome.ApplyStyle(estiloParafrafo.Name);
            // endereco.ApplyStyle(estiloParafrafo.Name);
            // valor.ApplyStyle(estiloParafrafo.Name);
            // dataDaCompra.ApplyStyle(estiloParafrafo.Name);

            documentImpressao.SaveToFile(@"Saida\dadosUser.docx", FileFormat.Docx);
        }
    }
}
