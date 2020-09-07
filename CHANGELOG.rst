Changelog
=========


v0.3.1 (2020-09-07)
-------------------
- Bump to 0.3.1. [Robin Thomas]
- Consolidated CONDITIONALS.rst into README.rst. [Robin Thomas]
- Let setup.cfg handle long_description and append conditionals doc.
  [Robin Thomas]
- Better short desc. [Robin Thomas]
- Added PyPy and CPython implementation classifications to setup.py.
  [Robin Thomas]
- Remove unused _wrap_as_standalone_function duplicate. [Robin Thomas]
- Indicate PyPy and PyPy3 support in README. (PyPy3 Travis build
  stumbles on Pandas install problems; my local PyPy3 environment (which
  required special NumPy source install with OpenBLAS config) shows a
  successful test suite. [Robin Thomas]
- Remove pypy3 travis target until pandas install problems can be fixed.
  [Robin Thomas]


v0.3.0 (2020-08-14)
-------------------
- Bump to version 0.3.0. [Robin Thomas]
- Include pypy and pypy3 in travis builds. [Robin Thomas]
- Add "batch updater" object (#21) [Robin Thomas]

  * Added batch capability to all formatting functions as well as format_with_dataframe.
  Minimal test coverage.

  * use "del listobj[:]" for 2.7 compatbility

  * Additional batch-updater tests; added batch updater docs to README.


v0.2.5 (2020-07-17)
-------------------
- Bump to version 0.2.5. [Robin Thomas]
- Fixes #20: BooleanCondition objects returned by API endpoints may lack
  a 'values' field instead of having a present 'values' field with an
  empty list of values. Allow for this in BooleanCondition constructor.
  Test coverage added for round-trip test of Boolean. [Robin Thomas]
- Argh no 3.9-dev yet. [Robin Thomas]
- Corrected version reference in sphinx docs. [Robin Thomas]
- Removed 3.6, added 3.9-dev to travis build` [Robin Thomas]
- Make collections.abc import 3.9-compatible. [Robin Thomas]
- Use full version string in sphnix docs. [Robin Thomas]
- Add docs badge to README. [Robin Thomas]
- Fix title in index.rst. [Robin Thomas]
- Try adding conditionals rst to docs. [Robin Thomas]
- Preserve original conditional rules for effective replacement of rules
  in one API call. [Robin Thomas]
- Add downloads badge. [Robin Thomas]


v0.2.4 (2020-05-04)
-------------------
- Bump to v0.2.4. [Robin Thomas]
- Make new Color.fromHex() and toHex() 2.7-compatible. [Robin Thomas]


v0.2.3 (2020-05-04)
-------------------
- Bump to v0.2.3. [Robin Thomas]
- Color model import and export as hex color (#17) [Sam Korn]

  * Add toHex function to Color model

  * tohex and fromhex functions for Color model

  * Use classmethod for hexstring constructor

  * tests for hex colors, additional checks for malformed hex inputs
- Results of check-manifest added to MANIFEST.in. [Robin Thomas]


v0.2.2 (2020-04-19)
-------------------
- Bump to v0.2.2. [Robin Thomas]
- Add MANIFEST.in to add VERSION file to sdist. [Robin Thomas]


v0.2.1 (2020-04-02)
-------------------
- Bump to v0.2.1. [Robin Thomas]
- Added support in DataFrame formatting for MultiIndex, either as index
  or as the columns object of the DataFrame. [Robin Thomas]
- Added docs/ to start sphinx autodoc generation. [Robin Thomas]
- Add wheel dep for bdist_wheel support. [Robin Thomas]


v0.2.0 (2020-03-31)
-------------------
- Bump to v0.2.0. [Robin Thomas]
- Fixes #10 (support setting row height or column width). [Robin Thomas]
- Added unbounded col and row ranges in format_cell_ranges test to
  ensure that formatting calls (not just _range_to_gridrange_object)
  succeed. [Robin Thomas]


v0.1.1 (2020-02-28)
-------------------
- Bump to v0.1.1. [Robin Thomas]
- Bare column row 14 (#15) [Robin Thomas]

  Fixes #14 -- support range strings that are unbounded on row dimension
  or column dimenstion.
- Oops typo. [Robin Thomas]
- Improve README intro and conditional docs text; attempt to include all
  .rst in package so that PyPI and others can see the other doc files.
  [Robin Thomas]


v0.1.0 (2020-02-11)
-------------------
- Bump to 0.1.0 for conditional formatting rules release. [Robin Thomas]
- Added doc about rule mutation and save() [Robin Thomas]
- Added conditional format rules documentation. [Robin Thomas]
- Added tests on effective cell format after conditional format rules
  apply. [Robin Thomas]
- Py2.7 MutableSequence does not mixin clear() [Robin Thomas]
- Tightened up add/delete of cond format rules, testing deletion of
  multiple rules. [Robin Thomas]
- Forbid illegal BooleanCondition.type values for data validation and
  conditional formatting ,respectively. [Robin Thomas]
- Realized that collections.abc is hoisted into collections module for
  backward compatibility already. [Robin Thomas]
- Add 2-3 compat for collections abc imports. [Robin Thomas]
- Final draft of conditional formatting implementation; test added,
  tests pass. Documentation not yet written. [Robin Thomas]
- Update README.rst. [Robin Thomas]


v0.0.9 (2020-02-09)
-------------------
- Bump to 0.0.9. [Robin Thomas]
- Data validation and prerequesites for conditional formatting 8 (#13)
  [Robin Thomas]

  * objects for conditional formatting added to data model

  * Implements data-validation feature requested in robin900/gspread-formatting#8.

  Test coverage included.

  * added GridRange object to models, ConditionalFormatRule class.

  * factored test code to allow Travis-style ssecret injection

  * merged in v0.0.8 changes from master; added full documentation for data validation;
  conditional format rules have all models in place, but no functions and no
  documentation in README.

  * add travis yml!

  * added requirements-test.txt so we can hopefully run tests in Travis

  * 2-3 compatible StringIO import in test

  * encrypt secrets files rather than env var approach to credentials and config

  * try encrypted files again

  * tighten up py versions in travis

  * make .tar.gz for travis secrets

  * bundle up secrets for travis ci

  * 2.7 compatible config reading

  * try a pip cache

  * fewer py builds


v0.0.8 (2020-02-06)
-------------------
- Fixes #12. Adds support for ColorStyle and all fields in which this
  object is now expected in the Sheets v4 API. See the Python or C# API
  documentation for reference, since the main REST API documentation
  still lacks mention of ColorStyle. [Robin Thomas]


v0.0.7 (2019-08-20)
-------------------
- Fixed setup.py problem that missed package contents. [Robin Thomas]
- Merge branch 'master' of github.com:robin900/gspread-formatting.
  [Robin Thomas]
- Update issue templates. [Robin Thomas]

  Added bug report template
- Bump to 0.0.7. [Robin Thomas]
- Add gspread-dataframe as dev req. [Robin Thomas]


v0.0.6 (2019-04-30)
-------------------
- Handle from_props cases where a format component is an empty dict of
  properties, so that comparing format objects round-trip works as
  expected, and so that format objects are as sparse as possible. [Robin
  Thomas]


v0.0.5 (2019-04-30)
-------------------
- Bump to 0.0.5. [Robin Thomas]
- Merge pull request #5 from robin900/fix-issue-4. [Robin Thomas]

  Conversion of API response's CellFormat properties failed for
- Conversion of API response's CellFormat properties failed for certain
  nested format components such as borders.bottom. Added test coverage
  to trigger bug, and code changes to solve the bug. Also added support
  of deprecated width= attribute for Border format component. [Robin
  Thomas]

  Fixes #4.


v0.0.4 (2019-03-26)
-------------------
- Bump VERSION to 0.0.4. [Robin Thomas]
- Merge pull request #2 from robin900/rthomas-dataframe-formatting.
  [Robin Thomas]

  Rthomas dataframe formatting
- Added docs and tests. [Robin Thomas]
- Working dataframe formatting, with test in test suite. Lacks complete
  documentation. [Robin Thomas]
- Added date-format test in response to user email; test confirms that
  package is working as expected. [Robin Thomas]
- Clean up of test suite, and provided instructions for dev and testing
  in README. [Robin Thomas]


v0.0.3 (2018-08-24)
-------------------
- Bump to 0.0.3, which fixes issue #1. [Robin Thomas]
- Fixed reference problem with NumberFormat.TYPES and Border.STYLES.
  [Robin Thomas]
- Added pypi badge. [Robin Thomas]
- Added format_cell_ranges, plus tests and documentation. [Robin Thomas]


v0.0.2 (2018-07-23)
-------------------
- Added get/set for frozen row and column counts. Bumped release to
  0.0.2. [Robin Thomas]


v0.0.1 (2018-07-20)
-------------------
- Tests pass; ready for version 0.0.1. [Robin Thomas]
- Initial commit. [Robin Thomas]


