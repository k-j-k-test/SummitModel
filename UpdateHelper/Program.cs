using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using Newtonsoft.Json;

class Program
{
    static void Main(string[] args)
    {
        //string from = @"C:\Users\wjdrh\OneDrive\Desktop\Test\SummitModel\ActuLight\bin\Debug\net48\UpdateFiles";
        //string to = @"C:\Users\wjdrh\OneDrive\Desktop\Test\SummitModel\ActuLight\bin\Debug\net48\";
        //args = new string[] { from, to };

        if (args.Length != 2)
        {
            Console.WriteLine("Usage: UpdateHelper.exe <source_path> <destination_path>");
            return;
        }

        string sourcePath = args[0].Trim('"');
        string destinationPath = args[1].Trim('"');
        string versionFilePath = Path.Combine(sourcePath, "version.json");
        string zipFilePath = Path.Combine(destinationPath, "update.zip");

        try
        {
            // 기존 ActuLight 프로세스 종료
            TerminateActuLightProcesses();

            string jsonContent = File.ReadAllText(versionFilePath);
            var versionInfo = JsonConvert.DeserializeObject<VersionInfo>(jsonContent);

            foreach (string file in versionInfo.UpdateFiles)
            {
                string sourceFile = Path.Combine(sourcePath, file);
                string destFile = Path.Combine(destinationPath, file);
                File.Copy(sourceFile, destFile, true);
            }

            Console.WriteLine("Update completed successfully.");

            // 임시 파일 및 폴더 삭제
            if (File.Exists(zipFilePath))
                File.Delete(zipFilePath);

            string mainAppPath = Path.Combine(destinationPath, "ActuLight.exe");
            if (File.Exists(mainAppPath))
            {
                Process.Start(mainAppPath);
            }
            else
            {
                throw new FileNotFoundException($"Main application not found at {mainAppPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Update failed. Error details:");
            Console.WriteLine(ex.ToString());
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    static void TerminateActuLightProcesses()
    {
        const int maxWaitTime = 10000; // 최대 10초 대기
        const int checkInterval = 500; // 0.5초마다 확인
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.ElapsedMilliseconds < maxWaitTime)
        {
            var processes = Process.GetProcessesByName("ActuLight");
            if (processes.Length == 0)
            {
                Console.WriteLine("All ActuLight processes have been terminated.");
                return;
            }

            foreach (var process in processes)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                        Console.WriteLine($"Terminated ActuLight process (ID: {process.Id})");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to terminate ActuLight process (ID: {process.Id}). Error: {ex.Message}");
                }
            }

            Thread.Sleep(checkInterval);
        }

        Console.WriteLine("Warning: Some ActuLight processes may still be running.");
    }

    class VersionInfo
    {
        public string Version { get; set; }
        public List<string> UpdateFiles { get; set; }
    }
}