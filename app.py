locale.Error: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:
File "/mount/src/enedis-streamlit/app.py", line 171, in <module>
    df_heatmap["Jour_semaine"] = df_heatmap["Horodate"].dt.day_name(locale="fr_FR")
                                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.13/site-packages/pandas/core/accessor.py", line 112, in f
    return self._delegate_method(name, *args, **kwargs)
           ~~~~~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.13/site-packages/pandas/core/indexes/accessors.py", line 132, in _delegate_method
    result = method(*args, **kwargs)
File "/home/adminuser/venv/lib/python3.13/site-packages/pandas/core/indexes/extension.py", line 95, in method
    result = attr(self._data, *args, **kwargs)
File "/home/adminuser/venv/lib/python3.13/site-packages/pandas/core/arrays/datetimes.py", line 1371, in day_name
    result = fields.get_date_name_field(
        values, "day_name", locale=locale, reso=self._creso
    )
File "pandas/_libs/tslibs/fields.pyx", line 165, in pandas._libs.tslibs.fields.get_date_name_field
File "pandas/_libs/tslibs/fields.pyx", line 644, in pandas._libs.tslibs.fields._get_locale_names
File "/usr/local/lib/python3.13/contextlib.py", line 141, in __enter__
    return next(self.gen)
File "/home/adminuser/venv/lib/python3.13/site-packages/pandas/_config/localization.py", line 47, in set_locale
    locale.setlocale(lc_var, new_locale)
    ~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^
File "/usr/local/lib/python3.13/locale.py", line 615, in setlocale
    return _setlocale(category, locale)
