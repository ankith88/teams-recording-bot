using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using TeamsBot.Core.Models;

namespace TeamsBot.Core.Interfaces
{
    public class ExportedFileInfo
    {
        public string Format { get; set; } = string.Empty; // docx, srt, json, txt
        public string FileName { get; set; } = string.Empty;
        public byte[] Content { get; set; } = System.Array.Empty<byte>();
    }

    public interface IStorageService
    {
        /// <summary>
        /// Sanitizes meeting subject into a valid operating system file name.
        /// Strips invalid OS characters: \ / : * ? " < > |
        /// </summary>
        string SanitizeFileName(string subject);

        /// <summary>
        /// Generates a non-clobbering file name if target file already exists.
        /// Example: "Weekly Sync.docx" -> "Weekly Sync (1).docx"
        /// </summary>
        string ResolveUniqueFilePath(string directoryPath, string baseFileName, string extension);

        /// <summary>
        /// Generates formatted export files (.docx, .srt, .json, .txt) for a meeting session.
        /// </summary>
        Task<List<ExportedFileInfo>> GenerateTranscriptExportsAsync(MeetingSession session);
    }
}
