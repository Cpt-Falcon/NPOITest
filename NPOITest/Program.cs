// See https://aka.ms/new-console-template for more information


using NPOI.XWPF.UserModel;

string[] templates = Directory.GetFiles(Directory.GetCurrentDirectory() + "\\TestDocs\\");
string working = templates.FirstOrDefault(x => Path.GetFileNameWithoutExtension(x) == "WorkingDocExample");
string linkBreak = templates.FirstOrDefault(x => Path.GetFileNameWithoutExtension(x) == "TestLinkBreak");
await GenerateNPOIDoc("working", working);
await GenerateNPOIDoc("TestLinkBreak", linkBreak);


async Task GenerateNPOIDoc(string fileName, string inputFile)
{
    byte[] bytes = await File.ReadAllBytesAsync(inputFile);
    using MemoryStream ms = new(bytes);
    XWPFDocument doc = new(ms);

    MemoryStream result = new();
    doc.Write(result);
    result.Flush();
    result.Position = 0;

    await using FileStream fs = new($"{fileName}.docx", FileMode.Create, FileAccess.Write);
    fs.Write(result.ToArray());
}