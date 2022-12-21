{
  'targets': [
    {
      'target_name': 'unispreadsheet',
      'sources': [ 'src/spreadsheet.cc' ],
      'link_settings': {
        'library_dirs': ['../include'],
        'libraries': [
          '-l:libspreadsheet.so'
        ],
        'ldflags': [
          # Ensure runtime linking is relative to sharp.node
          '-Wl,-s -Wl,--disable-new-dtags -Wl,-rpath=\'$$ORIGIN/../../include\''
        ]
      }
    }
  ]
}