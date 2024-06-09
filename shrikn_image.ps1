Add-Type -AssemblyName System.Drawing

# �摜�t�@�C���̊g���q�ꗗ��z��Ɋi�[���Ď擾
$imageExtensions = @("*.jpg", "*.jpeg", "*.png", "*.bmp")
$rootFolderPath = "C:\hogehoge\"

# �w�肵���t�H���_���̂��ׂẴT�u�t�H���_�����[�v
Get-ChildItem -Directory -Path $rootFolderPath -Recurse | ForEach-Object {
    $folderPath = $_.FullName

    # �e�T�u�t�H���_���̉摜�t�@�C�������[�v
    Get-ChildItem -File -Path $folderPath | ForEach-Object {
        # �������̃t�@�C������\��
        Write-Host $_

        if($imageExtensions -notcontains [System.IO.Path]::GetExtension($_.FullName)){
            Write-Host "�t�@�C��" $_ "�͉摜�t�@�C���łȂ����ߑΏۊO�ł��B"
            return
        }

        $image = [System.Drawing.Image]::FromFile($_.FullName)

        try {
            # �摜���^�e��(Portrait)�ł��邩�𔻕ʂ���ϐ� isPortrait
            $isPortrait = $image.Height -gt $image.Width

            # �V�����T�C�Y���v�Z
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
                Write-Host "�k���s�v"
                # �k�����s�v�ȏꍇ�A�X�L�b�v
                $image.Dispose()
                return
            }

            $newImage = New-Object System.Drawing.Bitmap($newWidth, $newHeight)

            # ���i���O���t�B�b�N�X�I�u�W�F�N�g���쐬
            $graphics = [System.Drawing.Graphics]::FromImage($newImage)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

            $graphics.DrawImage($image, 0, 0, $newWidth, $newHeight)
            $image.Dispose()

            # ���T�C�Y�摜���㏑���ۑ�
            $newImage.Save($_.FullName, $image.RawFormat)

            # ���\�[�X���
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