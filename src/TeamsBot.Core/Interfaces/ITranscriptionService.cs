using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using TeamsBot.Core.Models;

namespace TeamsBot.Core.Interfaces
{
    public interface ITranscriptionService
    {
        string EngineName { get; }
        
        /// <summary>
        /// Transcribes an audio stream with per-channel speaker attribution.
        /// </summary>
        Task<List<TranscriptSegment>> TranscribeAudioAsync(
            Stream audioStream, 
            Dictionary<uint, string> speakerChannelMap, 
            CancellationToken cancellationToken = default);

        /// <summary>
        /// Transcribes a single speaker audio segment in real time.
        /// </summary>
        Task<TranscriptSegment> TranscribeSegmentAsync(
            byte[] pcmData, 
            string speakerName, 
            TimeSpan startOffset, 
            CancellationToken cancellationToken = default);
    }
}
