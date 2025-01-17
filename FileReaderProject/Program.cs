using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;
using System.Diagnostics;

class Program
{
    static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        Console.WriteLine("Lütfen dosya yolunu girin:"); // Örn: C:\Users\ALİ\Desktop\example_file.txt
        string filePath = Console.ReadLine();

        // Dosya kontrolü.
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Hata: Belirtilen dosya yolu geçerli değil.");
            return;
        }

        // Dosya türü kontrolü.
        string fileExtension = Path.GetExtension(filePath).ToLower();

        if (fileExtension != ".txt" && fileExtension != ".docx" && fileExtension != ".pdf")
        {
            Console.WriteLine("Hata: Yalnızca .txt, .docx ve .pdf dosya türleri destekleniyor.");
            return;
        }

        // Dosya içeriğini okuma.
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

        // İçeriği analiz etme.
        AnalyzeContent(content);
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

    // İçerik analiz yöntemi.
    static void AnalyzeContent(string content)
    {
        // Kelimeleri bulma.
        string[] words = Regex.Split(content.ToLower(), @"\W+").Where(w => !string.IsNullOrEmpty(w)).ToArray();

        // Farklı kelime sayısı.
        int uniqueWordCount = words.Distinct().Count();

        // Tekrar eden kelimeler
        var repeatedWords = words.GroupBy(w => w)
                                 .Where(g => g.Count() > 1)
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
