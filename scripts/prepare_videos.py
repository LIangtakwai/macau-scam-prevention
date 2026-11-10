"""将提供的 MOV 转为浏览器兼容的 H.264/AAC MP4 并生成封面。"""
from pathlib import Path
import subprocess
import sys
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.tools'))
import imageio_ffmpeg
ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
files = sorted((ROOT / '材料/宣传视频').glob('*.mov'))
out = ROOT / 'materials/videos'
out.mkdir(exist_ok=True)
manifest = []
for i, source in enumerate(files, 1):
    target = out / f'video-{i}.mp4'
    print(f'Converting {i}/4: {source.name}', flush=True)
    subprocess.run([ffmpeg, '-hide_banner', '-loglevel', 'error', '-y', '-i', str(source),
        '-map', '0:v:0', '-map', '0:a?', '-vf', "scale='min(1280,iw)':-2",
        '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '24', '-pix_fmt', 'yuv420p',
        '-c:a', 'aac', '-b:a', '128k', '-movflags', '+faststart', str(target)], check=True)
    subprocess.run([ffmpeg, '-hide_banner', '-loglevel', 'error', '-y', '-ss', '1', '-i', str(target),
        '-frames:v', '1', '-update', '1', str(out / f'poster-{i}.jpg')], check=True)
    manifest.append(f'video-{i}.mp4 ← 材料/宣传视频/{source.name}')
(out / '素材对应.txt').write_text('\n'.join(manifest), encoding='utf-8')
print('All videos ready', flush=True)
