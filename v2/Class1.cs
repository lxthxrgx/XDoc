using System;
using System.Reflection.Metadata;
using Microsoft.Extensions.Primitives;

namespace XDoc.v2;

public sealed class PathToSave
{
    private string _pathToSave { get; set;}
    private string _fileName { get; set;}

    private PathToSave(string pathToSave, string fileName)
    {
        _pathToSave = pathToSave; 
        _fileName = fileName;
    }

    public static PathToSave Create(string filePath, string fileName)
    {
        if(string.IsNullOrEmpty(filePath))
            Console.WriteLine("file path cant be null");

        if(string.IsNullOrEmpty(fileName))
            Console.WriteLine("file name cant be null");

        return  new PathToSave(filePath, fileName);
    }

    public string Get()
    {
        return $"{_pathToSave}{_fileName}";
    }
}

public class Class1
{
    public required string _pathToFile {get; set;}
    public PathToSave _pathToSave {get; private set;}
    
    public static Class1 Start( string pathToFile, string pathToSave, string fileName)
    {
        if(string.IsNullOrEmpty(pathToFile))
            Console.WriteLine("Path to file can not be empty");

        if(!File.Exists(pathToFile))
            Console.WriteLine($"File not found. Check that the path is correct: {pathToFile}");
        
        Class1 newXDoc = new Class1
        {
            _pathToFile = pathToFile,
            _pathToSave = PathToSave.Create(pathToSave, fileName)
        };

        return newXDoc;
    }


}
