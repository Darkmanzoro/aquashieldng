15.0.1 ( Date : 29 September 2021 )
-----------------------------------

- Initial release

15.0.2 ( Date : 09 February 2022 )
----------------------------------

[Add] Xls Report with Credit, Debit total and all.

18.0.1.0.0 ( Date : 20 August 2026 )
------------------------------------

[Mig] Migrated from Odoo 15.0 to Odoo 18.0.
[Upd] Analytic column now reads ``analytic_distribution`` (``analytic_account_id``
      was removed from ``account.move.line`` in Odoo 16) and prints every
      analytic account with its percentage.  The PDF and the Xls export share
      one helper, ``account.move.line._sh_analytic_display()``, so both show
      identical text.
[Upd] Replaced the external ``html2text`` dependency with ``odoo.tools.html2plaintext``.
[Fix] Xls export no longer fails on journal entry names containing ``/`` -
      worksheet names are sanitised, truncated to 31 characters and de-duplicated.
[Fix] A draft entry has no name in Odoo 18 (it was ``/`` in 15), which xlwt
      wrote into the "Journal Entry #" cell as ``0``; it prints ``/`` again.
[Upd] Report template uses ``t-out`` instead of the deprecated ``t-esc``.
