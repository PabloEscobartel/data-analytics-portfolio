# Legacy and Superseded Scripts

This folder contains older or experimental scripts that are retained for traceability but are not part of the current final pipeline.

## Contents

The folder includes:

- earlier entity-matcher versions;
- earlier regression and diagnostics scripts;
- earlier GJRM/bivariate-probit scripts;
- small one-off utilities and API tests.

## Current Replacements

- Entity matching: use `train NLP/bond_entity_matcher_v10.py`.
- Final regressions: use `analysis_pipeline_final/scripts/run_regressions_upsize_batch.py`.
- Final diagnostics: use `analysis_pipeline_final/scripts/run_diagnostics_upsize_batch.py`.
- Final GJRM: use `analysis_pipeline_final/scripts/run_biprobit_gjrm_cluster_bootstrap_v4.R`.

Do not use scripts from this folder for final results without reviewing them first.
