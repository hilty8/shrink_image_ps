# アセンブリの読み込み
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# 画像ファイルの縮小処理を実行するフォルダパス
$folderPath = "C:\hoge"

#-----------------------------------------------------------
# 画像ファイルのサイズを変更
#-----------------------------------------------------------

# 画像ファイルの拡張子一覧を配列に格納して取得
$imageExtensions = @("*.jpg", "*.jpeg", "*.png", "*.bmp")

foreach ($extension in $imageExtensions) {
    Get-ChildItem -File -Path $folderPath -Filter $extension -Recurse | ForEach-Object {
        Write-Host $_.FullName

        $image = [System.Drawing.Image]::FromFile($_.FullName)
        try {
            # 画像がタテ長(Portrait)であるかを判別する変数 isPortrait
            $isPortrait = $image.Height -gt $image.Width

            # 画像がタテ長、かつ横幅が800pxより大きい場合
            if ($isPortrait -and $image.Width -gt 800) {
                $ratio = $image.Height / $image.Width
                $newWidth = 800
                $newHeight = [math]::Round($newWidth * $ratio)
            }
            # 画像がヨコ長、かつ横幅が1200pxより大きい場合
            elseif (-not $isPortrait -and $image.Width -gt 1200) {
                $ratio = $image.Height / $image.Width
                $newWidth = 1200
                $newHeight = [math]::Round($newWidth * $ratio)
            }
            #-------------------------------
            # 画像サイズを変更して上書き保存
            #-------------------------------
            $canvas = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
            if($canvas){
                $graphics = [System.Drawing.Graphics]::FromImage($canvas)
                $graphics.DrawImage($image, (New-Object System.Drawing.Rectangle(0, 0, $canvas.Width, $canvas.Height)))
                $image.Dispose()
                # Resizeデータの保存
                $canvas.Save($_.FullName, [System.Drawing.Imaging.ImageFormat]::Jpeg)

                $graphics.Dispose()
                $canvas.Dispose()

            }else{
                echo "変数 canvas が存在しません。"
            }
        }

        finally {
            $image.Dispose()
        }
    }
}
