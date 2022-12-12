{
  'targets': [
    {
      'target_name': 'unispreadsheet',
      'sources': [ 'src/spreadsheet.cc' ],
      'cflags!': [ '-fno-exceptions' ],
      'cflags_cc!': [ '-fno-exceptions' ],
      'libraries': ['<!(pwd)/include/libspreadsheet.so'],
    }
  ]
}