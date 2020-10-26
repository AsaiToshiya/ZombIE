# ZombIE

ZombIE は、決して死なないゾンビのような Internet Explorer のためのユーティリティー クラスで、ファイルのダウンロードを自動化します。


## 使い方

1. [ZombIE.cs](ZombIE.cs) をプロジェクトにコピーします。
2. `UIAutomationClient` と `UIAutomationTypes` への参照を追加します。
3. `ZombIE` の名前空間をプロジェクトの名前空間に変更します。
4. Internet Explorer を操作して、ダウンロードのダイアログまたは通知バーが表示されたタイミングで、`ZombIE` の `DownloadFileTo` メソッドを呼び出します。

例:

```csharp
var driver = new InternetExplorerDriver();
// ...
driver.FindElement(By.Id("download")).Click();
ZombIE.DownloadFileTo(@"c:\temp\file.doc");
```


## 動作環境

 * Internet Explorer 11
 * .NET Framework >= 4.0
