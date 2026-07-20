# Selecting referees and matches

Selections are parsed by the command-line interface using a compact syntax.

## Examples

```text
1        # first referee or match
1,3,5    # selected items
1-4      # range
4-2      # interpreted as 2-4
         # blank means all
0        # all in multi-selection menus
T        # todos / all
```

Invalid tokens are ignored. If no valid index remains, the analysis is not launched.

## Batch strategies

- **One referee:** useful for quality control and iterative inspection.
- **Several referees:** useful when a known subset should be processed.
- **All referees:** useful for reproducible batch generation after source validation.
