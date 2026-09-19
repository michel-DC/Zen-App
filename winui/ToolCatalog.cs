using System.Collections.Generic;

namespace Zen.WinUI;

internal sealed record ToolDefinition(
    string Title,
    string Description,
    string Extension,
    string[] Extensions,
    bool ExtractsText = false);

internal static class ToolCatalog
{
    internal static IReadOnlyDictionary<string, ToolDefinition> All { get; } =
        new Dictionary<string, ToolDefinition>
        {
            ["docx_pdf"] = new("DOCX en PDF", "Convertir un document Word au format PDF.", "pdf", [".docx"]),
            ["pdf_docx"] = new("PDF en DOCX", "Créer un document Word à partir d’un PDF.", "docx", [".pdf"]),
            ["image_jpg"] = new("Image en JPG", "Convertir une image PNG, BMP, GIF ou TIFF au format JPG.", "jpg", [".png", ".bmp", ".gif", ".tif", ".tiff"]),
            ["png_pdf"] = new("PNG en PDF", "Créer un PDF à partir d’une image PNG.", "pdf", [".png"]),
            ["jpg_pdf"] = new("JPG en PDF", "Créer un PDF à partir d’une image JPG.", "pdf", [".jpg", ".jpeg"]),
            ["pdf_text"] = new("Texte d’un PDF", "Lire et copier le texte d’un PDF directement dans Zen.", "", [".pdf"], true),
            ["docx_text"] = new("Texte d’un DOCX", "Lire et copier le texte d’un document Word directement dans Zen.", "", [".docx"], true),
            ["image_ocr"] = new("Texte d’une image", "Reconnaître et copier le texte d’une image avec le moteur OCR local.", "", [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"], true),
            ["crop"] = new("Rogner une image", "Recadrer une zone précise de l’image.", "png", [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"]),
            ["rounded"] = new("Arrondir les coins", "Ajouter des coins arrondis sur fond transparent.", "png", [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"]),
            ["circle"] = new("Créer une image circulaire", "Découper l’image au centre dans un cercle transparent.", "png", [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"]),
        };

    internal static string GetActionLabel(string id) => id switch
    {
        "docx_pdf" => "Convertir en PDF",
        "pdf_docx" => "Convertir en DOCX",
        "png_pdf" or "jpg_pdf" => "Créer le PDF",
        "image_jpg" => "Créer le JPG",
        "pdf_text" or "docx_text" or "image_ocr" => "Extraire le texte",
        "crop" => "Rogner l’image",
        "rounded" => "Arrondir les coins",
        "circle" => "Créer l’image circulaire",
        _ => "Traiter le fichier",
    };
}
