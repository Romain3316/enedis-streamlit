from streamlit.testing.v1 import AppTest

SCRIPT = '''
import streamlit as st
from export_ui import prepare_export
st.text_input('Company', key='company_name')
st.text_area('Note', key='note_context')
st.checkbox('Preview', key='preview')
def build(name):
    st.session_state['_builds'] = st.session_state.get('_builds', 0) + 1
    return (name + st.session_state.company_name + st.session_state.note_context).encode()
a = prepare_export('a', 'Prepare A', lambda: build('a'))
b = prepare_export('b', 'Prepare B', lambda: build('b'))
st.session_state['_outputs'] = (a, b)
'''


def test_exports_are_lazy_reused_invalidated_and_session_local():
    at = AppTest.from_string(SCRIPT).run()
    assert not at.exception
    assert at.session_state['_outputs'] == (None, None)
    at.button(key='prepare_a').click().run()
    assert at.session_state['_prepared_export']['data'] == b'a'
    at.checkbox(key='preview').check().run()
    assert at.session_state['_builds'] == 1
    assert at.session_state['_outputs'] == (b'a', None)
    at.text_area(key='note_context').set_value('new note').run()
    assert at.session_state['_outputs'] == (None, None)
    at.button(key='prepare_b').click().run()
    assert at.session_state['_prepared_export']['data'] == b'bnew note'
    at.run()
    assert at.session_state['_outputs'] == (None, b'bnew note')
    at.session_state['_source_files'] = {'curve': {'name':'second.csv', 'content':b'new meter'}}
    at.run()
    assert at.session_state['_outputs'] == (None, None)
    other = AppTest.from_string(SCRIPT).run()
    assert other.session_state['_outputs'] == (None, None)
    assert not at.exception and not other.exception
