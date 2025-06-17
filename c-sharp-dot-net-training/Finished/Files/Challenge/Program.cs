// variables to store number of Microsoft file types
int wordType = 0;
int excelType = 0;
int powerPointType = 0;
// variable to store total number of files
int totalFiles;

// variables to store the file size of each Microsoft file types
long wordSize = 0;
long excelSize = 0;
long powerPointSize = 0;
// total size
long totalSize;

// get the current path of this directory
string curPath = Directory.GetCurrentDirectory();

// get the FileCollection directory
string fileCollPath = Path.Combine(curPath, "FileCollection");

// enumerate each file in the FileCollection directory
List<string> thefiles = new List<string>(Directory.EnumerateFiles(fileCollPath));
foreach (string file in thefiles)
{
    // get the file info for each file
    FileInfo fileInfo = new FileInfo(file);
    if (fileInfo.Exists)
    {
        // check which type of file (word, excel, or ppt)
        if (fileInfo.Extension == ".docx")
        {
            wordType++;
            wordSize += fileInfo.Length;
        }
        else if (fileInfo.Extension == ".xlsx")
        {
            excelType++;
            excelSize += fileInfo.Length;
        }
        else if (fileInfo.Extension == ".pptx")
        {
            powerPointType++;
            powerPointSize += fileInfo.Length;
        }
        else
        {
            continue;
        }
    }
}

// calaculate the totals
totalFiles = wordType + excelType + powerPointType;
totalSize = wordSize + excelSize + powerPointSize;

// display the results
Console.WriteLine("--- Results ---");
Console.WriteLine($"Total Files: {totalFiles}");
Console.WriteLine($"Word Count: {wordType}");
Console.WriteLine($"Excel Count: {excelType}");
Console.WriteLine($"PowerPoint Count: {powerPointType}");
Console.WriteLine("---");
Console.WriteLine($"Total Size: {totalSize:N0}");
Console.WriteLine($"Word Size: {wordSize:N0}");
Console.WriteLine($"Excel Size: {excelSize:N0}");
Console.WriteLine($"PowerPoint Size: {powerPointSize:N0}");