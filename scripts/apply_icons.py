from pathlib import Path
import shutil
root=Path(__file__).resolve().parents[1]
icons=root/'assets'/'icons'
ios=root/'ios'/'App'/'App'/'Assets.xcassets'/'AppIcon.appiconset'
if ios.exists():
    for name in ('AppIcon-512@2x.png','Contents.json'):
        shutil.copy2(icons/'ios'/name,ios/name)
    print('Ícone oficial IP aplicado ao iOS')
android=root/'android'/'app'/'src'/'main'/'res'
if android.exists():
    for source in (icons/'android').glob('mipmap-*'):
        target=android/source.name; target.mkdir(exist_ok=True,parents=True)
        for file in source.iterdir():shutil.copy2(file,target/file.name)
    print('Ícone oficial IP aplicado ao Android')
