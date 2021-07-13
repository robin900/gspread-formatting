Changelog
=========


v1.0.3 (2021-07-13)
-------------------
- Bump to v1.0.3. [Robin Thomas]
- Fixes #34. Cells with no formatting or data validation rules were
  causing KeyError exceptions in get_effective_format() and similar
  functions. These functions now properly return None without raising an
  exception. [Robin Thomas]


v1.0.2 (2021-07-01)
-------------------
- Bump to v1.0.2. [Robin Thomas]
- Compare model classes to json schema from discovery URI, using new
  script; remove foregroundColorStyle from CellFormat class as it's not
  in json schema. (Border.width is in json schema but is deprecated and
  thus Border model class is correctly coded.) [Robin Thomas]


v1.0.1 (2021-06-30)
-------------------
- Bump to v1.0.1. [Robin Thomas]
- Fixes #33 -- 'link' property of TextFormat now supported. [Robin
  Thomas]


v1.0.0 (2021-05-13)
-------------------
- Bump to v1.0.0. [Robin Thomas]
- Fix for #31 (#32) [Robin Thomas]

  Fixes #31. Allows for Sheets API's tendency to include empty objects
  for default color values in API responses.
- No longer CI with pypy. [Robin Thomas]
- Revert "Attempt to constrain use of old rsa pkg to 2.7 CI build, and
  also" [Robin Thomas]

  This reverts commit 5cc83c67036ba5d004de997b613a34e9a8550f24.
- Revert "Attempt to constrain use of old rsa pkg to 2.7 CI build, and
  also" [Robin Thomas]

  This reverts commit 5cc83c67036ba5d004de997b613a34e9a8550f24.
- Revert "Try TravisCI conditional use of rsa<=4.1 again" [Robin Thomas]

  This reverts commit b8ecaad65f0876385a1585237d756ee1fd450fb0.
- Oops use rsa<4.1. [Robin Thomas]
- Try TravisCI conditional use of rsa<=4.1 again. [Robin Thomas]
- Attempt to constrain use of old rsa pkg to 2.7 CI build, and also
  avoid the github dependabot alert. [Robin Thomas]
- Pin rsa to < 4.1 so that Python 2.7 CI can still run. [Robin Thomas]
- Added paranoid test of absent sheetId in GridRange props, to prevent
  accidental regression. [Robin Thomas]
- Improved, more concise code for Color.fromHex and .toHex(). [Robin
  Thomas]
- Tighten up travis install. [Robin Thomas]
- Try explicit directory caching to make pip cache work as expected for
  pandas wheel. [Robin Thomas]
- Add 3.9 to travis. [Robin Thomas]
- Pin six to >=1.12.0 in travis to avoid weird environmental dependency
  problem. [Robin Thomas]
- Move to travis-ci.com. [Robin Thomas]


v0.3.7 (2020-11-23)
-------------------
- Bump to v0.3.7. [Robin Thomas]
- Corrected error in conditional format rules example code in README.
  [Robin Thomas]
- Fixed typo in README. [Robin Thomas]
- Fixed typos in batch call documentation. [Robin Thomas]


v0.3.6 (2020-11-12)
-------------------
- Bump to v0.3.6. [Robin Thomas]
- Allow for absent sheetId property in GridRange objects coming from API
  (suspected abrupt change in Sheets API behavior!) [Robin Thomas]
- Added extra example for clearing data validation rule with None.
  [Robin Thomas]


v0.3.5 (2020-11-10)
-------------------
- Bump to v0.3.5. [Robin Thomas]
- Fixes #26. Allows `None` as rule parameter to
  set_data_validation_rule* functions, which will clear data validation
  rule for the relevant cells. [Robin Thomas]


v0.3.4 (2020-10-22)
-------------------
- Bump to v0.3.4. [Robin Thomas]
- More informative exception message when BooleanCondition receives non-
  list/tuple for values parameter. [Robin Thomas]
- Increased already-high test coverage. [Robin Thomas]
- Removed dead link to now-inlined conditional formatting doc. [Robin
  Thomas]
- Correct doc/sphinx annoyances. [Robin Thomas]


v0.3.3 (2020-09-24)
-------------------
- Bump to version v0.3.3. [Robin Thomas]
- Fixes #24. [Robin Thomas]

  A certain set of functions that exist both in batch and standalone mode
  are dynamically bound as local names in the functions subpackage. That makes
  them undiscoverable by IDEs like PyCharm. Adding a straightforward import
  statement for these function names -- even though the names are re-bound
  immediately with wrapped standalone versions of the functions -- makes
  the function names visible to PyCharm.


v0.3.2 (2020-09-16)
-------------------
- Bump to v0.3.2. [Robin Thomas]
- Fixes #23. Test coverage added. [Robin Thomas]
- Support InterpolationPoint.colorStyle. [Robin Thomas]


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


