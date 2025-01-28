using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;
using System.Diagnostics;
using System.Windows.Forms;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string filePath = SelectFileWithDialog();
        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Hata: Geçerli bir dosya seçilmedi.");
            return;
        }

        if (!File.Exists(filePath))
        {
            Console.WriteLine("Hata: Belirtilen dosya yolu geçerli değil.");
            return;
        }

        string fileExtension = Path.GetExtension(filePath).ToLower();

        if (fileExtension != ".txt" && fileExtension != ".docx" && fileExtension != ".pdf")
        {
            Console.WriteLine("Hata: Yalnızca .txt, .docx ve .pdf dosya türleri destekleniyor.");
            return;
        }

        string content = string.Empty;
        try
        {
            if (fileExtension == ".txt")
            {
                content = File.ReadAllText(filePath);
            }
            else if (fileExtension == ".docx")
            {
                content = ReadDocxContent(filePath);
            }
            else if (fileExtension == ".pdf")
            {
                content = ReadPdfContent(filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Hata: Dosya okunurken bir hata oluştu. {ex.Message}");
            return;
        }

        AnalyzeContent(content);
    }

    static string SelectFileWithDialog()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Title = "Lütfen bir dosya seçin";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
        }
        return null;
    }

    // .docx dosyalarını okumak için yardımcı yöntem.
    static string ReadDocxContent(string filePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            return doc.MainDocumentPart.Document.Body.InnerText;
        }
    }

    // .pdf dosyalarını okumak için yardımcı yöntem.
    static string ReadPdfContent(string filePath)
    {
        string content = string.Empty;
        using (PdfDocument pdf = PdfDocument.Open(filePath))
        {
            foreach (var page in pdf.GetPages())
            {
                content += page.Text;
            }
        }
        return content;
    }

    static readonly HashSet<string> Conjunctions = new HashSet<string>
    {
     "ve","ile","ama","fakat","lakin","çünkü","yalnız","ya","veya","1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30","a","b","c","ç","d","e","f","g","ğ","h","ı","i","j","k","l","m","n","o","ö","p","r","s","ş","t","u","ü","v","y","z","q","w","x",
     "ki","oysa","dahi","de","da","hem","ise","bu","en","bir","emco","was","here",
    };

    // İçerik analiz yöntemi.
    static void AnalyzeContent(string content)
    {
        // Kelimeleri bulma.
        string[] words = Regex.Split(content.ToLower(), @"\W+").Where(w => !string.IsNullOrEmpty(w)).ToArray();


        var filteredWords = words.Where(word => !Conjunctions.Contains(word)).ToArray();

        // Farklı kelime sayısı.
        int uniqueWordCount = filteredWords.Distinct().Count();

        // Tekrar eden kelimeler.
        var repeatedWords = filteredWords.GroupBy(w => w)
                                         .Where(g => g.Count() > 1)
                                          .OrderByDescending(g => g.Count())
                                         .ToDictionary(g => g.Key, g => g.Count());

        // Noktalama işaretlerini bulma.
        var punctuationMarks = Regex.Matches(content, @"[.,;:!?()""'“”‘’]")
                                    .Cast<Match>()
                                    .Select(m => m.Value)
                                    .GroupBy(p => p)
                                    .ToDictionary(g => g.Key, g => g.Count());

        // Sonuçları yazdırma.
        Console.WriteLine($"Toplam farklı kelime sayısı: {uniqueWordCount}");

        Console.WriteLine("\nTekrar eden kelimeler ve tekrar sayıları:");
        foreach (var word in repeatedWords)
        {
            Console.WriteLine($"- {word.Key}: {word.Value} kez");
        }

        Console.WriteLine("\nKullanılan noktalama işaretleri:");
        foreach (var punctuation in punctuationMarks)
        {
            Console.WriteLine($"- '{punctuation.Key}': {punctuation.Value} kez");
        }
    }
}
