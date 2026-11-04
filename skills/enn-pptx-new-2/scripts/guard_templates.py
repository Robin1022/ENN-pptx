#!/usr/bin/env python3
"""Explicit template hash snapshot/verification; not an execution hook."""
import argparse
import hashlib
import json
from pathlib import Path
import sys


def template_hashes(directory):
    if not directory.is_dir():
        raise ValueError(f"Template directory missing: {directory}")
    files = sorted(p for p in directory.rglob('*')
                   if p.is_file() and p.suffix.lower() in {'.pptx', '.potx'})
    if not files:
        raise ValueError(f"No templates found: {directory}")
    return {str(p.relative_to(directory)): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in files}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--snapshot', type=Path, metavar='MANIFEST')
    mode.add_argument('--verify', type=Path, metavar='MANIFEST')
    parser.add_argument('--templates', type=Path,
                        default=Path(__file__).resolve().parents[1] / 'templates')
    args = parser.parse_args()
    directory = args.templates.resolve()
    try:
        current = template_hashes(directory)
        if args.snapshot:
            target = args.snapshot.resolve()
            if target == directory or directory in target.parents:
                raise ValueError('Manifest must be outside the template directory')
            with target.open('x', encoding='utf-8') as stream:
                json.dump({'version': 1, 'directory': str(directory), 'sha256': current},
                          stream, ensure_ascii=False, indent=2)
            print(f'Recorded {len(current)} template hashes: {target}')
        else:
            data = json.loads(args.verify.read_text(encoding='utf-8'))
            if not isinstance(data, dict) or data.get('version') != 1 or data.get('directory') != str(directory):
                raise ValueError('Manifest version or template directory mismatch')
            expected = data.get('sha256')
            if not isinstance(expected, dict) or not expected:
                raise ValueError('Invalid or empty baseline')
            changed = sorted(name for name in set(expected) | set(current)
                             if expected.get(name) != current.get(name))
            if changed:
                raise ValueError('Templates changed, missing or added: ' + ', '.join(changed))
            print(f'PASS: {len(current)} templates unchanged')
    except (OSError, ValueError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
