using System;
using System.Data;
using System.IO;
using System.Text;
using ExcelDataReader;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        // 엑셀 파일 경로
        string filePath = "./DataBaseSetRes/exdata.xlsx";

        // 하위 디렉토리 이름
        string subdirectoryName = "DataBaseSetRes";

        // 현재 프로젝트 디렉토리 경로
        string projectDirectory = Directory.GetCurrentDirectory();

        // 하위 디렉토리 경로 생성
        string subdirectoryPath = Path.Combine(projectDirectory, subdirectoryName);

        // 하위 디렉토리가 없으면 생성
        if (!Directory.Exists(subdirectoryPath))
        {
            Directory.CreateDirectory(subdirectoryPath);
        }

        // 파일 스트림으로 엑셀 파일 열기
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            // IExcelDataReader 인터페이스로 데이터 읽기
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                // 첫 번째 시트 선택
                reader.Read();
                
                // 열 이름 출력
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    Console.Write(reader.GetValue(i) + "\t");
                }
                Console.WriteLine();

                // 행 데이터 출력
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        Console.Write(reader.GetValue(i) + "\t");
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}