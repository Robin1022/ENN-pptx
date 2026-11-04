#!/usr/bin/env python3
"""Validate the shared skill, sync Codex metadata and build a plugin ZIP."""
import argparse
import json
from pathlib import Path
import re
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def build(output_directory):
    manifest = json.loads((ROOT / 'plugin.json').read_text(encoding='utf-8'))
    name, version = manifest['name'], manifest['version']
    if not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', name):
        raise ValueError('Invalid plugin name')
    if not re.fullmatch(r'\d+\.\d+\.\d+', version):
        raise ValueError('Use a three-part release version, such as 1.0.0')
    skill = ROOT / 'skills' / name
    required = ['SKILL.md', 'CODEX.md', 'agents/openai.yaml',
                'scripts/guard_templates.py', 'scripts/office/validate.py',
                'scripts/check_delivery.py', 'scripts/check_layout.py',
                'templates/company-template.pptx', 'templates/campus-hire-defence.pptx']
    for rel in required:
        if not (skill / rel).is_file():
            raise ValueError(f'Missing skill resource: {rel}')
    if not re.search(r'^name:\s*' + re.escape(name) + r'\s*$',
                     (skill / 'SKILL.md').read_text(encoding='utf-8'), re.MULTILINE):
        raise ValueError('Skill name and plugin name differ')
    market = json.loads((ROOT / '.agents/plugins/marketplace.json').read_text(encoding='utf-8'))
    entries = market['plugins']
    if len(entries) != 1 or entries[0]['name'] != name or entries[0]['source'] != {'source': 'local', 'path': './'}:
        raise ValueError('Marketplace must point to this plugin root')
    # Root manifest is the source of truth; old Codex clients read this overlay.
    compatibility = {k: v for k, v in manifest.items() if k not in ('$schema', 'extensions')}
    compatibility.update(manifest['extensions']['com.openai'])
    compatibility['skills'] = './skills/'
    target = ROOT / '.codex-plugin/plugin.json'
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(compatibility, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    files = [ROOT / 'plugin.json', target, ROOT / '.agents/plugins/marketplace.json', ROOT / 'README.md']
    files.extend(p for p in skill.rglob('*') if p.is_file()
                 and '__pycache__' not in p.parts and p.suffix != '.pyc' and p.name != '.DS_Store')
    output_directory = output_directory.resolve()
    if output_directory == skill or skill in output_directory.parents:
        raise ValueError('Archive output must be outside the skill directory')
    output_directory.mkdir(parents=True, exist_ok=True)
    archive = output_directory / f'{name}-{version}.zip'
    with zipfile.ZipFile(archive, 'w', compression=zipfile.ZIP_DEFLATED) as bundle:
        for path in sorted(files):
            if path.is_symlink():
                raise ValueError(f'Package must contain actual files: {path}')
            bundle.write(path, path.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(archive) as bundle:
        if bundle.testzip() is not None:
            raise ValueError('ZIP integrity check failed')
    print(f'Built {archive} ({len(files)} files)')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'dist')
    args = parser.parse_args()
    try:
        build(args.output_dir)
    except (OSError, ValueError, KeyError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
