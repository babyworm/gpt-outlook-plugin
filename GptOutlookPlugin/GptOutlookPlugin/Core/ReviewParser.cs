using System.Collections.Generic;
using System.Text.RegularExpressions;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public static class ReviewParser
    {
        // 유연한 패턴: 줄바꿈 변형, 공백, 대소문자 허용
        private static readonly Regex CorrectionBlock = new Regex(
            @"\[CORRECTION\]\s*\r?\n" +
            @"\s*ORIGINAL\s*:\s*(?<orig>.+?)\s*\r?\n" +
            @"\s*CORRECTED\s*:\s*(?<corr>.+?)\s*\r?\n" +
            @"\s*REASON\s*:\s*(?<reason>.+?)\s*\r?\n" +
            @"\s*\[/CORRECTION\]",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // 대안 패턴: 마크다운 스타일 (AI가 포맷을 약간 벗어날 때)
        private static readonly Regex AltBlock = new Regex(
            @"\*?\*?ORIGINAL\*?\*?\s*:\s*(?<orig>.+?)\s*\r?\n" +
            @"\s*\*?\*?CORRECTED\*?\*?\s*:\s*(?<corr>.+?)\s*\r?\n" +
            @"\s*\*?\*?REASON\*?\*?\s*:\s*(?<reason>.+?)(?:\r?\n|$)",
            RegexOptions.IgnoreCase | RegexOptions.Multiline);

        /// <summary>
        /// AI 응답에서 교정 블록을 파싱.
        /// [CORRECTION] 블록 먼저 시도, 실패 시 대안 패턴 시도.
        /// 블록이 없으면 null 반환 (일반 텍스트 응답으로 fallback).
        /// </summary>
        public static List<ReviewCorrection> Parse(string response)
        {
            if (string.IsNullOrEmpty(response)) return null;
            if (response.Contains("[NO_ISSUES]")) return new List<ReviewCorrection>();

            // 1차: [CORRECTION] 블록 패턴
            var matches = CorrectionBlock.Matches(response);
            if (matches.Count > 0)
                return ExtractFromMatches(matches);

            // 2차: ORIGINAL/CORRECTED/REASON 대안 패턴
            var altMatches = AltBlock.Matches(response);
            if (altMatches.Count > 0)
                return ExtractFromMatches(altMatches);

            return null;
        }

        private static List<ReviewCorrection> ExtractFromMatches(MatchCollection matches)
        {
            var result = new List<ReviewCorrection>();
            foreach (Match m in matches)
            {
                var orig = m.Groups["orig"].Value.Trim().Trim('"', '`');
                var corr = m.Groups["corr"].Value.Trim().Trim('"', '`');
                var reason = m.Groups["reason"].Value.Trim();

                if (!string.IsNullOrEmpty(orig) && !string.IsNullOrEmpty(corr))
                {
                    result.Add(new ReviewCorrection
                    {
                        Original = orig,
                        Corrected = corr,
                        Reason = reason
                    });
                }
            }
            return result.Count > 0 ? result : null;
        }
    }
}
