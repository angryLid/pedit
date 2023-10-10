export type Attr<T> = T extends string ? `@_${T}` : never;
