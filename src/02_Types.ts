/**
 * Project types
 */
type RowNumber = number;

enum Column {
    null = 0, A = 1,
    B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
}

interface SheetProtocol {
    sheet: GAS.Spreadsheet.Sheet;
    getLastRow() : RowNumber;
}