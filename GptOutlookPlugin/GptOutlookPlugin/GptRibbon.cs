using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin
{
    [ComVisible(true)]
    public class GptRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public event Action<FeatureMode> OnModeRequested;
        public event Action<FeatureMode, string> OnSelectionModeRequested;
        public event Action OnSettingsRequested;

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""GptTab"" label=""GPT Email Assistant"">
        <group id=""GptMainGroup"" label=""AI Actions"">
          <button id=""btnReview"" label=""Review"" imageMso=""ReviewTrackChanges"" size=""large""
                  onAction=""OnReviewClick"" screentip=""Review and proofread the current email""/>
          <button id=""btnCompose"" label=""Compose"" imageMso=""NewMailMessage"" size=""large""
                  onAction=""OnComposeClick"" screentip=""Write a new email with guidance""/>
          <button id=""btnAutoReply"" label=""Auto Reply"" imageMso=""Reply"" size=""large""
                  onAction=""OnAutoReplyClick"" screentip=""Auto-draft a reply to the current email""/>
          <button id=""btnTranslate"" label=""Translate"" imageMso=""SetLanguage"" size=""large""
                  onAction=""OnTranslateClick"" screentip=""Translate the current email""/>
          <button id=""btnSummarize"" label=""Summarize"" imageMso=""Consolidate"" size=""large""
                  onAction=""OnSummarizeClick"" screentip=""Summarize the current email""/>
        </group>
        <group id=""GptSettingsGroup"" label=""Settings"">
          <button id=""btnSettings"" label=""Settings"" imageMso=""PropertySheet"" size=""normal""
                  onAction=""OnSettingsClick"" screentip=""Configure AI provider and preferences""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnReviewClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Review);
        }

        public void OnComposeClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Compose);
        }

        public void OnAutoReplyClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.AutoReply);
        }

        public void OnTranslateClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Translate);
        }

        public void OnSummarizeClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Summarize);
        }

        public void OnSettingsClick(IRibbonControl control)
        {
            OnSettingsRequested?.Invoke();
        }

        public void OnContextReviewClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Review, selection);
        }

        public void OnContextTranslateClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Translate, selection);
        }

        public void OnContextComposeClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Compose, selection);
        }
    }
}
