from pathlib import Path

def _repo_root() -> Path:
    # Retorna o diretório raiz do repositório (um nível acima deste script)
    return Path(__file__).resolve().parents[1]

def main() -> int:
    root = _repo_root()
    icons_dir = root / "assets" / "images" / "icons"
    src = icons_dir / "icon.png"

    if not src.exists():
        alt = root / "assets" / "icons" / "icon.png"
        if alt.exists(): src = alt

    if not src.exists():
        raise FileNotFoundError(f"Não encontrei o ícone fonte em: {icons_dir / 'icon.png'}")

    icons_dir.mkdir(parents=True, exist_ok=True)
    from PIL import Image
    base = Image.open(src).convert("RGBA")
    sizes = [16, 24, 32, 48, 64, 128, 256]
    for s in sizes:
        out_png = icons_dir / f"icon_{s}.png"
        resized = base.resize((s, s), Image.LANCZOS)
        resized.save(out_png, format="PNG")

    out_ico = icons_dir / "Habibot.ico"
    base.save(out_ico, format="ICO", sizes=[(s, s) for s in sizes])
    (icons_dir / "icon.png").write_bytes(src.read_bytes())
    print(f"OK: gerado {out_ico}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
