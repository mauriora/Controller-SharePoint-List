// Alternatively:

// export type IfEquals<X, Y, A, B> =
//     [2] & [0, 1, X] extends [2] & [0, 1, Y] & [0, infer W, unknown]
//     ? W extends 1 ? B : A
//     : B;

export type IfEquals<X, Y, A, B> =
    (<T>() => T extends X ? 1 : 2) extends
    (<T>() => T extends Y ? 1 : 2) ? A : B;

export type WritableKeysOf<T> = {
    [P in keyof T]: IfEquals<{ [Q in P]: T[P] }, { -readonly [Q in P]: T[P] }, P, never>
}[keyof T];

export type WritablePart<T> = Pick<T, WritableKeysOf<T> & string>;
