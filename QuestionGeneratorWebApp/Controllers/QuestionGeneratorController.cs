using Microsoft.AspNetCore.Mvc;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace QuestionGeneratorWebApp.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class QuestionGeneratorController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<QuestionGeneratorController> _logger;

        public QuestionGeneratorController(IWebHostEnvironment env, ILogger<QuestionGeneratorController> logger)
        {
            _env = env;
            _logger = logger;
        }

        [HttpPost("generate-mcq")]
        public IActionResult GenerateMcq([FromBody] McqGenerationRequest request)
        {
            try
            {
                var outputPath = Path.Combine(_env.ContentRootPath, "Generated");
                Directory.CreateDirectory(outputPath);

                var createdFiles = McqQuestionGenerator.McqGenerateQuestionDoc(
                    outputPath,
                    request.QuestionNumber,
                    request.SubjectList,
                    request.SequenceList,
                    request.IncludeAnswerTags,
                    request.MultiSet,
                    request.SetCount
                );

                if (createdFiles.Count > 0)
                {
                    return Ok(new { success = true, files = createdFiles, message = "MCQ document(s) generated successfully!" });
                }

                return BadRequest(new { success = false, message = "Failed to generate document(s)." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating MCQ document");
                return StatusCode(500, new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpPost("generate-saq")]
        public IActionResult GenerateSaq([FromBody] SaqGenerationRequest request)
        {
            try
            {
                var outputPath = Path.Combine(_env.ContentRootPath, "Generated");
                Directory.CreateDirectory(outputPath);

                var files = SaqQuestionGenerator.CreateSaqQuestion(
                    outputPath,
                    request.QuestionNumber,
                    request.QuestionMark,
                    request.SubjectListBangla,
                    request.SubjectListEnglish,
                    request.SequenceList,
                    request.MultiSet,
                    request.SetCount
                );

                if (files.Count > 0)
                {
                    var generatedFiles = files.Take(10).ToList();
                    return Ok(new { success = true, files = generatedFiles, message = "SAQ documents generated successfully!" });
                }

                return BadRequest(new { success = false, message = "Failed to generate documents." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating SAQ document");
                return StatusCode(500, new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpPost("add-answer-tag")]
        public async Task<IActionResult> AddAnswerTag(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest(new { success = false, message = "No file uploaded." });

                var uploadsPath = Path.Combine(_env.ContentRootPath, "Uploads");
                Directory.CreateDirectory(uploadsPath);

                var filePath = Path.Combine(uploadsPath, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                AddAnswerTagToDoc(filePath);

                var fileName = Path.GetFileName(filePath);
                return Ok(new { success = true, fileName = fileName, message = "Answer tags added successfully!" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding answer tags");
                return StatusCode(500, new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpPost("add-subject-name")]
        public async Task<IActionResult> AddSubjectName(IFormFile file, [FromForm] string subjectList, [FromForm] string sequenceList)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest(new { success = false, message = "No file uploaded." });

                var uploadsPath = Path.Combine(_env.ContentRootPath, "Uploads");
                Directory.CreateDirectory(uploadsPath);

                var filePath = Path.Combine(uploadsPath, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                AddSubjectNameToDoc(filePath, subjectList, sequenceList);

                var fileName = Path.GetFileName(filePath);
                return Ok(new { success = true, fileName = fileName, message = "Subject names added successfully!" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding subject names");
                return StatusCode(500, new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpGet("download/{fileName}")]
        public IActionResult Download(string fileName)
        {
            try
            {
                var generatedPath = Path.Combine(_env.ContentRootPath, "Generated", fileName);
                var uploadsPath = Path.Combine(_env.ContentRootPath, "Uploads", fileName);

                string filePath = null;
                if (System.IO.File.Exists(generatedPath))
                    filePath = generatedPath;
                else if (System.IO.File.Exists(uploadsPath))
                    filePath = uploadsPath;

                if (filePath == null)
                    return NotFound(new { success = false, message = "File not found." });

                var bytes = System.IO.File.ReadAllBytes(filePath);
                return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading file");
                return StatusCode(500, new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // Helper methods using Open XML instead of Office Interop
        private void AddAnswerTagToDoc(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var tables = body.Descendants<Table>().ToList();

                // Map option letters to target row index (0-based)
                var optionRowMap = new Dictionary<string, int> { { "a", 2 }, { "b", 3 }, { "c", 4 }, { "d", 5 } };

                foreach (var table in tables)
                {
                    try
                    {
                        var rows = table.Descendants<TableRow>().ToList();
                        if (rows.Count < 8) continue;

                        // 8th row, 1st column contains a/b/c/d
                        var ansCell = rows[7].Descendants<TableCell>().FirstOrDefault();
                        if (ansCell == null) continue;
                        string ansText = string.Join("", ansCell.Descendants<Text>().Select(t => t.Text)).Trim().ToLowerInvariant();
                        if (!optionRowMap.TryGetValue(ansText, out int targetRowIndex)) continue;
                        if (targetRowIndex >= rows.Count) continue;

                        var targetRow = rows[targetRowIndex];
                        var cells = targetRow.Descendants<TableCell>().ToList();
                        for (int col = 0; col < Math.Min(2, cells.Count); col++)
                        {
                            var cell = cells[col];
                            // Skip if cell is blank
                            string originalClean = string.Join("", cell.Descendants<Text>().Select(t => t.Text)).Trim();
                            if (string.IsNullOrWhiteSpace(originalClean)) continue;

                            // Replace text with "Answer"
                            cell.RemoveAllChildren<Paragraph>();
                            var p = new Paragraph(new Run(new Text("Answer")));
                            cell.AppendChild(p);
                        }
                    }
                    catch { continue; }
                }

                doc.MainDocumentPart.Document.Save();
            }
        }

        private void AddSubjectNameToDoc(string filePath, string subjectList, string sequenceList)
        {
            var subjects = subjectList
                            .Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .ToList();

            var sequences = sequenceList
                            .Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .ToList();

            if (subjects.Count != sequences.Count)
                throw new Exception("subjects and sequences must have equal counts!");

            using (var doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var tables = body.Descendants<Table>().ToList();

                int tableIndex = 0;
                foreach (var table in tables)
                {
                    tableIndex++;

                    for (int i = 0; i < subjects.Count; i++)
                    {
                        string subject = subjects[i];
                        string seq = sequences[i];

                        var parts = seq.Split('-', StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length != 2 ||
                            !int.TryParse(parts[0], out int start) ||
                            !int.TryParse(parts[1], out int end))
                        {
                            continue;
                        }

                        if (tableIndex >= start && tableIndex <= end)
                        {
                            try
                            {
                                var rows = table.Descendants<TableRow>().ToList();
                                if (rows.Count == 0) continue;
                                var firstRow = rows[0];
                                var cells = firstRow.Descendants<TableCell>().ToList();
                                if (cells.Count < 2) continue;

                                var subjectCell = cells[1];
                                subjectCell.RemoveAllChildren<Paragraph>();
                                subjectCell.AppendChild(new Paragraph(new Run(new Text(subject))));
                            }
                            catch { }
                        }
                    }
                }

                doc.MainDocumentPart.Document.Save();
            }
        }
    }

    public class McqGenerationRequest
    {
        public int QuestionNumber { get; set; } = 100;
        public string SubjectList { get; set; } = "phy, chem, math, bio";
        public string SequenceList { get; set; } = "1-25,26-50,51-75,76-100";
        public bool IncludeAnswerTags { get; set; } = false;
        public bool MultiSet { get; set; } = false;
        public int SetCount { get; set; } = 1;
    }

    public class SaqGenerationRequest
    {
        public int QuestionNumber { get; set; } = 100;
        public int QuestionMark { get; set; } = 2;
        public string SubjectListBangla { get; set; } = "পদার্থবিজ্ঞান, রসায়ন";
        public string SubjectListEnglish { get; set; } = "Phy,Chem";
        public string SequenceList { get; set; } = "1-50,51-100";
        public bool MultiSet { get; set; } = false;
        public int SetCount { get; set; } = 1;
    }
}
