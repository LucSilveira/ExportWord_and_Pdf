using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace Exportando_arquivos
{
    class Program
    {
        static void Main(string[] args)
        {
            // Criando um novo documento com o nome: documento
            Document documentoExercicio = new Document();

            // Criando uma seção dentro do documento
            // A cada seção criada uma nova página é adicionada;
            Section secaoCapa = documentoExercicio.AddSection();

            // Criando um titulo na primeira página
            Paragraph titulo = secaoCapa.AddParagraph();

            // Inserindo o conteúdo que ira aparecer no titulo criado
            titulo.AppendText("Titulo inicia\n\n");

            Section secaoTexto = documentoExercicio.AddSection();

            Paragraph textoExcercicio  = secaoTexto.AddParagraph();

            textoExcercicio.AppendText("Morto não fala\nMorto não vê\nSe reclamar do meu sitema\nO próximo morto é você\n<3<3<3");

            // Centralizando o titulo no centre do documento
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            //  Instanciando a classe ParagraphStyle dentro do sistema
            ParagraphStyle estiloParafrafo = new ParagraphStyle(documentoExercicio);

            // Definindo o nome da classe estiloParagraph
            estiloParafrafo.Name = "Cor do titulo";

            // Pinta a propriedade TextColor de AzulEscuro
            estiloParafrafo.CharacterFormat.TextColor = Color.DarkBlue;

            // Definindo que o elemento é true para bold
            // pois o mesmo é um atributo no sistema
            estiloParafrafo.CharacterFormat.Bold = true;

            // Adicionar o estilo e colocar como usavel no nosso documento
            documentoExercicio.Styles.Add(estiloParafrafo);

            titulo.ApplyStyle(estiloParafrafo.Name);

            documentoExercicio.SaveToFile(@"Saida\Exercicio.docx", FileFormat.Docx);
        }
    }
}
