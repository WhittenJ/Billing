import cx_Freeze

executables = [cx_Freeze.Executable('C:\\Users\\jwhitten\\PycharmProjects\\Billing\\EzNetv2.0.py')]

cx_Freeze.setup(
    name='EZ Net Form Filler',
    options={'build_exe': {'packages': ['selenium']}},
    executables=executables
)
