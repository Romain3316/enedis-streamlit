"""Generate one requested export per session; discard it when inputs change."""
import hashlib
import pickle

import streamlit as st
from dossier_schema import SETTINGS


def prepare_export(key, label, build, extra=None):
    state = st.session_state
    context = (
        {name: state.get(name, definition.get('default')) for name, definition in SETTINGS.items()},
        state.get('_source_files'), state.get('_solar_snapshot'),
        state.get('_resolved_location'),
    )
    revision = hashlib.sha256(pickle.dumps(context, protocol=5)).hexdigest()
    retained = state.get('_prepared_export')
    if retained and (retained['revision'] != revision or (retained['key'] == key and retained['extra'] != extra)):
        state.pop('_prepared_export', None)
        retained = None
    if st.button(label, key='prepare_' + key, use_container_width=True):
        # Release the previous output before generating another potentially large one.
        state.pop('_prepared_export', None)
        retained = None
        try:
            log_memory('before_' + key)
            with st.spinner('Préparation du fichier…'):
                data = build()
            log_memory('after_' + key)
            retained = {'key': key, 'revision': revision, 'extra': extra, 'data': data}
            state['_prepared_export'] = retained
        except Exception as exc:
            st.error(f"Le fichier n'a pas pu être préparé : {exc}")
    if retained and retained['key'] == key:
        return retained['data']
    return None


def log_memory(stage):
    """Log container memory counters only, never dossier contents or identities."""
    from pathlib import Path
    import json
    counters = {'stage': stage}
    try:
        import resource
        import sys
        scale = 1024 ** 2 if sys.platform == 'darwin' else 1024
        counters['process_peak_mib'] = round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / scale, 1)
    except ImportError:
        pass
    for name in ('memory.current', 'memory.max', 'memory.events'):
        try:
            counters[name] = (Path('/sys/fs/cgroup') / name).read_text().strip()
        except OSError:
            pass
    print('[memory] ' + json.dumps(counters), flush=True)
