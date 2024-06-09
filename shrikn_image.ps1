Add-Type -AssemblyName System.Drawing

# 画像ファイルの拡張子一覧を配列に格納して取得
$imageExtensions = @("*.jpg", "*.jpeg", "*.png", "*.bmp")
$rootFolderPath = "C:\hogehoge\"

# 指定したフォルダ内のすべてのサブフォルダをループ
Get-ChildItem -Directory -Path $rootFolderPath -Recurse | ForEach-Object {
    $folderPath = $_.FullName

    # 各サブフォルダ内の画像ファイルをループ
    Get-ChildItem -File -Path $folderPath | ForEach-Object {
        # 処理中のファイル名を表示
        Write-Host $_

        if($imageExtensions -notcontains [System.IO.Path]::GetExtension($_.FullName)){
            Write-Host "ファイル" $_ "は画像ファイルでないため対象外です。"
            return
        }

        $image = [System.Drawing.Image]::FromFile($_.FullName)

        try {
            # 画像がタテ長(Portrait)であるかを判別する変数 isPortrait
            $isPortrait = $image.Height -gt $image.Width

            # 新しいサイズを計算
            if ($isPortrait -and $image.Width -gt 800) {
                $ratio = $image.Height / $image.Width
                $newWidth = 800
                $newHeight = [math]::Round($newWidth * $ratio)
            }
            elseif (-not $isPortrait -and $image.Width -gt 1200) {
                $ratio = $image.Height / $image.Width
                $newWidth = 1200
                $newHeight = [math]::Round($newWidth * $ratio)
            }
            else {
                Write-Host "縮小不要"
                # 縮小が不要な場合、スキップ
                $image.Dispose()
                return
            }

            $newImage = New-Object System.Drawing.Bitmap($newWidth, $newHeight)

            # 高品質グラフィックスオブジェクトを作成
            $graphics = [System.Drawing.Graphics]::FromImage($newImage)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

            $graphics.DrawImage($image, 0, 0, $newWidth, $newHeight)
            $image.Dispose()

            # リサイズ画像を上書き保存
            $newImage.Save($_.FullName, $image.RawFormat)

            # リソース解放
            $graphics.Dispose()
            $newImage.Dispose()
        }
        finally {
            if ($image -ne $null) {
                $image.Dispose()
            }
        }
    }
}