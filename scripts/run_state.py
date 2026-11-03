"""Record codex-ppt-style production evidence. Stdlib only; does not author slides."""
import argparse
import hashlib
import json
from pathlib import Path
from datetime import datetime, timezone
import zipfile
import xml.etree.ElementTree as ET

GATES = ('outline', 'style', 'backend', 'sample')

def stamp():
    return datetime.now(timezone.utc).isoformat()

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def existing(value):
    p = Path(value).resolve()
    if not p.is_file():
        raise ValueError(f'Missing evidence file: {p}')
    return {'path': str(p), 'sha256': sha(p)}

def save(path, state):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding='utf-8')
    tmp.replace(path)

def verify_file(evidence):
    p = Path(evidence['path'])
    if not p.is_file() or sha(p) != evidence['sha256']:
        raise ValueError(f'Evidence missing or changed: {p}')

def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    for cmd in ['init', 'gate', 'plan', 'start', 'record', 'verify']:
        q = sub.add_parser(cmd)
        q.add_argument('--project', required=True)
        if cmd == 'init':
            q.add_argument('--pages', type=int, required=True)
            q.add_argument('--route', choices=['editable', 'image'], default='editable')
        elif cmd == 'gate':
            q.add_argument('--name', choices=GATES, required=True)
            q.add_argument('--basis', required=True)
            q.add_argument('--evidence')
            q.add_argument('--waived', action='store_true')
        elif cmd == 'plan':
            q.add_argument('--spec', required=True)
        elif cmd in ('start', 'record'):
            q.add_argument('--slide', required=True)
            if cmd == 'start':
                q.add_argument('--worker', required=True)
            else:
                q.add_argument('--artifact', required=True)
                q.add_argument('--preview', required=True)
                q.add_argument('--notes-ready', action='store_true')
                q.add_argument('--qa-note', required=True)
        elif cmd == 'verify':
            q.add_argument('--deck', required=True)
    args = parser.parse_args(argv)
    project = Path(args.project).resolve()
    statepath = project / '.build' / 'deck_state.json'
    if args.command == 'init':
        if statepath.exists():
            raise ValueError('Existing run state will not be overwritten; use a new project for a new run.')
        if args.pages < 1:
            raise ValueError('--pages must be positive')
        state = {'schema_version': 1, 'created': stamp(), 'pages': args.pages,
                 'route': args.route, 'gates': {}, 'jobs': {}}
    else:
        state = json.loads(statepath.read_text(encoding='utf-8'))
        if args.command == 'gate':
            if state['jobs']:
                raise ValueError('Cannot change gates after planning; begin a new run for a changed contract.')
            earlier = GATES[:GATES.index(args.name)]
            if any(k not in state['gates'] for k in earlier):
                raise ValueError(f'First record preceding gates: {earlier}')
            if not args.basis.strip():
                raise ValueError('A real approval or waiver basis is required')
            item = {'basis': args.basis, 'waived': args.waived, 'recorded': stamp()}
            if args.evidence:
                item['evidence'] = existing(args.evidence)
            if args.name in ('outline', 'sample') and not (args.evidence or args.waived):
                raise ValueError('Outline/sample approval needs evidence, or an explicit waiver basis')
            state['gates'][args.name] = item
        elif args.command == 'plan':
            if any(k not in state['gates'] for k in GATES):
                raise ValueError('Record outline, style, backend and sample gates before full production')
            if state['jobs']:
                raise ValueError('Jobs already exist; do not overwrite production history')
            specfile = existing(args.spec)
            spec = json.loads(Path(specfile['path']).read_text(encoding='utf-8'))
            if spec.get('route') != state['route']:
                raise ValueError('Spec route must match confirmed route')
            method = spec.get('method', {})
            if not method.get('tool') or not method.get('mode'):
                raise ValueError('Record actual tool and mode inherited from the sample')
            if state['route'] == 'editable' and method['mode'] != 'native':
                raise ValueError('Editable route requires native mode')
            if state['route'] == 'image' and method['mode'] not in ('generate', 'edit'):
                raise ValueError('Image route requires the confirmed image generate/edit method')
            pages = spec.get('slides', [])
            if len(pages) != state['pages']:
                raise ValueError('Spec page count differs from requested count')
            jobs = {}
            for number, slide in enumerate(pages, 1):
                sid = f'slide_{number:02d}'
                if slide.get('id') != sid or slide.get('order') != number:
                    raise ValueError(f'Spec must use contiguous actual page order: {sid}')
                for field in ('title', 'role', 'key_points', 'layout_intent', 'local_context', 'notes_outline'):
                    if not slide.get(field):
                        raise ValueError(f'{sid} missing {field}')
                if not isinstance(slide.get('source_assets'), list):
                    raise ValueError(f'{sid} needs source_assets, possibly []')
                assets = []
                for asset in slide['source_assets']:
                    if asset.get('role') not in ('evidence', 'style'):
                        raise ValueError(f'{sid}: asset role must distinguish evidence/style')
                    for field in ('path', 'locator', 'fidelity'):
                        if not asset.get(field):
                            raise ValueError(f'{sid}: asset missing {field}')
                    p = Path(asset['path'])
                    assets.append({**asset, **existing(p if p.is_absolute() else project / p)})
                jobs[sid] = {'status': 'pending', 'task': slide, 'source_assets': assets,
                             'canonical_terms': spec.get('canonical_terms', []), 'method': method}
            state.update({'spec': specfile, 'method': method, 'jobs': jobs})
            jobsdir = project / '.build' / 'jobs'
            for sid, job in jobs.items():
                save(jobsdir / f'{sid}.json', job)
        elif args.command in ('start', 'record'):
            if args.slide not in state['jobs']:
                raise ValueError(f'Unknown/unplanned slide: {args.slide}')
            job = state['jobs'][args.slide]
            for asset in job['source_assets']:
                verify_file(asset)
            if args.command == 'start':
                if job['status'] == 'dispatched':
                    raise ValueError('Job is already dispatched')
                job.update({'status': 'dispatched', 'worker': args.worker, 'started': stamp()})
                job.pop('result', None)
            else:
                if job['status'] != 'dispatched':
                    raise ValueError('Record start/worker before recording a result')
                if not args.qa_note.strip():
                    raise ValueError('Record an actual inspection result')
                job['result'] = {'artifact': existing(args.artifact), 'preview': existing(args.preview),
                                 'notes_ready': args.notes_ready, 'qa_note': args.qa_note, 'recorded': stamp()}
                job['status'] = 'recorded'
            save(project / '.build' / 'jobs' / f'{args.slide}.json', job)
            state.pop('verification', None)
        elif args.command == 'verify':
            if len(state['jobs']) != state['pages']:
                raise ValueError('Missing planned pages')
            verify_file(state['spec'])
            for gate in state['gates'].values():
                if 'evidence' in gate:
                    verify_file(gate['evidence'])
            for sid, job in state['jobs'].items():
                result = job.get('result', {})
                if job['status'] != 'recorded' or not result.get('notes_ready'):
                    raise ValueError(f'Unfinished page/notes: {sid}')
                for asset in job['source_assets']:
                    verify_file(asset)
                verify_file(result['artifact'])
                verify_file(result['preview'])
            deck = existing(args.deck)
            with zipfile.ZipFile(deck['path']) as z:
                doc = ET.fromstring(z.read('ppt/presentation.xml'))
                count = len(doc.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst/{http://schemas.openxmlformats.org/presentationml/2006/main}sldId'))
            if count != state['pages']:
                raise ValueError(f'Actual deck has {count} pages, expected {state["pages"]}')
            state['verification'] = {'deck': deck, 'recorded': stamp(),
                                     'claim': 'Run evidence and count only; not visual, editability, notes-content or scientific validation.'}
    save(statepath, state)
    print(json.dumps({'command': args.command, 'state': str(statepath),
                      'recorded_pages': sum(j['status'] == 'recorded' for j in state['jobs'].values()),
                      'total_pages': state['pages']}, ensure_ascii=False))

if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as error:
        raise SystemExit(f'ERROR: {error}')
