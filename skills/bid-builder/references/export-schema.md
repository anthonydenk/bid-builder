# CAD export schema

## Expected columns

The deterministic parser recognizes these headers, case-insensitively:

| Required | Header |
|---|---|
| yes | Type |
| yes | Item |
| yes | Quantity |
| yes | Description |
| no | Desired Profit |
| yes | Net Cost |
| yes | Unit Profit Amount |
| yes | Unit Price |
| yes | Extended Cost |
| yes | Extended Profit |
| yes | Net Price |
| no | Category |

The header row may appear after title/blank rows. The parser searches each sheet for a row containing the required headers.

## Arithmetic checks

For every non-empty line, validate within two cents:

- `Quantity × Net Cost = Extended Cost`
- `Quantity × Unit Profit Amount = Extended Profit`
- `Net Cost + Unit Profit Amount = Unit Price`
- `Extended Cost + Extended Profit = Net Price`

The source total is the sum of every parsed row's `Net Price`. Do not recalculate the commercial bid from only the rows shown to the customer.

## Public versus internal rows

The default public type is `Item`. Common internal types include `Freight`, `Labor`, `Overhead`, `Travel`, and `Textura`. Catalog rules override type defaults. This matters when the CAD system carries a real product inside an internal-looking row.

Blank template rows and rows whose quantity, prices, and description are all empty or zero are ignored.
