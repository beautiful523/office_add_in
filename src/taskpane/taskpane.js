/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("results").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const body = context.document.body;
    body.load("text");

    await context.sync();

    const paragraphs = body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      const firstParagraph = paragraphs.items[0];
      const wordText = firstParagraph.getText();
      await context.sync();

      const text = wordText.value;
      const words = text.split(" ");
      if (words.length >= 3) {
        const firstWordRange = firstParagraph.search(words[0], { matchCase: false, matchWholeWord: true });
        const secondWordRange = firstParagraph.search(words[1], { matchCase: false, matchWholeWord: true });
        const thirdWordRange = firstParagraph.search(words[2], { matchCase: false, matchWholeWord: true });

        firstWordRange.load("font/bold");
        secondWordRange.load("font/underline");
        thirdWordRange.load("font/size");

        await context.sync();
        document.getElementById("result-error").style.display = "none";
        document.getElementById("results").style.display = "block";
        document.getElementById("result1").innerText = firstWordRange.items[0].font.bold == true ? "是" : "否";
        document.getElementById("result2").innerText = secondWordRange.items[0].font.underline === "None" ? "否" : "是";
        document.getElementById("result3").innerText = thirdWordRange.items[0].font.size;
      } else {
        document.getElementById("result-error").style.display = "block";
        document.getElementById("results").style.display = "none";
        document.getElementById("result-error").innerText = " 文档词汇不够";
      }
    } else {
      document.getElementById("result-error").style.display = "block";
      document.getElementById("results").style.display = "none";
      document.getElementById("result-error").innerText = "文档中暂无段落";
    }
  });
}
