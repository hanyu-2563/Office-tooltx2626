//only for test   
//This program has not undergone rigorous experiments and is not yet available as a finished product
using Microsoft.Office.Interop.Word;
//ttpp
public void ConvertNumberedListToText()
{
    // 获取当前Word应用程序实例
    Application wordApp = Globals.ThisAddIn.Application;

    // 获取所选段落
    Paragraph selectedPara = wordApp.Selection.Range.Paragraphs[1];

    // 检查是否为数字列表
    if (selectedPara.Range.ListFormat.ListType == WdListType.wdListBullet
        || selectedPara.Range.ListFormat.ListType == WdListType.wdListNoNumbered
        || selectedPara.Range.ListFormat.ListType == WdListType.wdListOutlineNumbering)
    {
        // 将列表项的格式设置为文本格式
        selectedPara.Range.ListFormat.ApplyListTemplate(
            ListTemplate: wordApp.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates[3],
            ContinuePreviousList: false,
            ApplyTo: WdListApplyTo.wdListApplyToWholeList,
            DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);

        // 更新段落样式，确保正确呈现文本格式的列表项
        selectedPara.Range.Select();
        wordApp.Selection.ParagraphFormat.Update();
    }
}
