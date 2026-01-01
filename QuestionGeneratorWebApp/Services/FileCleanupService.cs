using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Hosting;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace QuestionGeneratorWebApp.Services
{
    public class FileCleanupService : BackgroundService
    {
        private readonly ILogger<FileCleanupService> _logger;
        private readonly IWebHostEnvironment _env;
        private readonly TimeSpan _cleanupInterval = TimeSpan.FromMinutes(1); // Run cleanup every minute
        private readonly TimeSpan _fileMaxAge = TimeSpan.FromMinutes(5); // Delete files older than 5 minutes

        public FileCleanupService(ILogger<FileCleanupService> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _env = env;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("File Cleanup Service started");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await Task.Delay(_cleanupInterval, stoppingToken);
                    CleanupOldFiles();
                }
                catch (OperationCanceledException)
                {
                    // Service is stopping
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error occurred during file cleanup");
                }
            }

            _logger.LogInformation("File Cleanup Service stopped");
        }

        private void CleanupOldFiles()
        {
            _logger.LogInformation("Starting file cleanup...");

            var generatedPath = Path.Combine(_env.ContentRootPath, "Generated");
            var uploadsPath = Path.Combine(_env.ContentRootPath, "Uploads");

            int deletedCount = 0;

            // Clean up Generated directory
            if (Directory.Exists(generatedPath))
            {
                deletedCount += CleanupDirectory(generatedPath);
            }

            // Clean up Uploads directory
            if (Directory.Exists(uploadsPath))
            {
                deletedCount += CleanupDirectory(uploadsPath);
            }

            _logger.LogInformation($"File cleanup completed. Deleted {deletedCount} file(s)");
        }

        private int CleanupDirectory(string directoryPath)
        {
            int deletedCount = 0;
            var cutoffTime = DateTime.Now - _fileMaxAge;

            try
            {
                var files = Directory.GetFiles(directoryPath);
                foreach (var file in files)
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.LastWriteTime < cutoffTime)
                        {
                            File.Delete(file);
                            deletedCount++;
                            _logger.LogDebug($"Deleted old file: {fileInfo.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"Failed to delete file: {file}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error cleaning up directory: {directoryPath}");
            }

            return deletedCount;
        }
    }
}
