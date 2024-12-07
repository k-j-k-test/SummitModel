using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

class Program
{
    static readonly string[] ExcludedFiles = new string[]
    {
        "recentFiles.json",
        "settings.json"
    };

    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: UpdateHelper.exe <source_path> <destination_path>");
            return;
        }

        string sourcePath = args[0].Trim('"');
        string destinationPath = args[1].Trim('"');
        string zipFilePath = Path.Combine(destinationPath, "update.zip");

        try
        {
            // 기존 ActuLight 프로세스 종료
            TerminateActuLightProcesses();

            // 모든 파일과 폴더 복사 (제외 항목 제외)
            CopyDirectory(sourcePath, destinationPath);

            Console.WriteLine("Update completed successfully.");

            // 임시 파일 삭제
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

    static void CopyDirectory(string sourceDir, string destDir)
    {
        // 대상 디렉토리가 없으면 생성
        if (!Directory.Exists(destDir))
        {
            Directory.CreateDirectory(destDir);
        }

        // 파일 복사 (제외 항목 제외)
        foreach (string file in Directory.GetFiles(sourceDir))
        {
            string fileName = Path.GetFileName(file);
            string extension = Path.GetExtension(file);

            // 제외할 파일 확인
            if (extension.Equals(".pdb", StringComparison.OrdinalIgnoreCase) ||
                ExcludedFiles.Contains(fileName, StringComparer.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Skipping excluded file: {fileName}");
                continue;
            }

            string destFile = Path.Combine(destDir, fileName);
            try
            {
                File.Copy(file, destFile, true);
                Console.WriteLine($"Copied file: {destFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to copy file {file}: {ex.Message}");
            }
        }

        // 하위 디렉토리 복사
        foreach (string dir in Directory.GetDirectories(sourceDir))
        {
            string destSubDir = Path.Combine(destDir, Path.GetFileName(dir));
            CopyDirectory(dir, destSubDir);
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
}