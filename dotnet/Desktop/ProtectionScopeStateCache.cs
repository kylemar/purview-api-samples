using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Formats.Asn1.AsnWriter;

namespace Purview_API_Explorer
{
    public enum CallPurviewType
    {
        Dont,
        Inline,
        Offline
    }

    static public class ProtectionScopeStateCache
    {
        public static CallPurviewType Prompts = CallPurviewType.Dont;
        public static CallPurviewType Responses = CallPurviewType.Dont;

        static public void ParseProtectionScopeState(dynamic parsedJson)
        {
             Prompts = CallPurviewType.Dont;
             Responses = CallPurviewType.Dont;

            try
            {
                JArray? scopes = parsedJson["value"] as JArray;
                if (scopes == null)
                {
                    return;
                }

                foreach (var scope in scopes)
                {
                    string executionMode = scope["executionMode"]?.ToString() ?? string.Empty;
                    string activities = scope["activities"]?.ToString() ?? string.Empty;
                    JArray? locations = scope["locations"] as JArray;
                    JArray? policyActions = scope["policyActions"] as JArray;

                    if (activities.Contains("uploadText"))
                    {
                        if (executionMode == "evaluateInline")
                        {
                            Prompts = CallPurviewType.Inline;
                        }
                        else if (executionMode == "evaluateOffline")
                        {
                            if (Prompts != CallPurviewType.Inline)
                            {
                                Prompts = CallPurviewType.Offline;
                            }
                        }
                    }

                    if (activities.Contains("downloadText"))
                    {
                        if (executionMode == "evaluateInline")
                        {
                            Responses = CallPurviewType.Inline;
                        }
                        else if (executionMode == "evaluateOffline")
                        {
                            if (Responses != CallPurviewType.Inline)
                            {
                                Responses = CallPurviewType.Offline;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error parsing ProtectionScopeState: {ex.Message}");
            }
        }
    }
}
