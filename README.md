# excel2img

Excelファイルを画像ファイルに変換します。  
<https://github.com/osawa-koki/img2excel>プロジェクトの兄弟プロジェクト。  

お遊びプロジェクト。  

![成果物](./docs/img/fruits.gif)  

## 使用した画像ファイル

- [タコ](https://frame-illust.com/?p=13667)
- [スズメ](https://frame-illust.com/?p=13680)
- [キツネ](https://frame-illust.com/?p=9584)

## 使い方

```shell
./excel2img.exe /t 対象画像ファイルパス /o 出力先ファイル名

# One By One
./excel2img.exe /t tako.xlsx /o tako.png
./excel2img.exe /t tanuki.xlsx /o tanuki.png
./excel2img.exe /t suzume.xlsx /o suzume.png
```
