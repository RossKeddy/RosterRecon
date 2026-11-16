# Contributing to RosterRecon

Bug reports, focused feature requests, documentation improvements, and pull
requests are welcome. Before submitting a change:

1. Run `python -m pip install --editable '.[dev]'` in a virtual environment.
2. Run `python -m pytest`.
3. Do not include credentials, session cookies, or collected employee data.
4. Keep collection changes conservative so they do not increase lockout or
   rate-limit risk without a clear operator-controlled option.

For larger changes, open an issue first so the design can be discussed.
