// The bundled Node entry point of MathJax 4 ships no type declarations;
// declare only the pieces of the API that we use.
declare module '@mathjax/src' {
  interface MathJaxNode {
    // opaque handle to the rendered node, passed back to the adaptor
    readonly __brand?: never;
  }
  const MathJax: {
    init(config: Record<string, unknown>): Promise<unknown>;
    tex2svgPromise(
      math: string,
      options?: {display?: boolean},
    ): Promise<MathJaxNode>;
    startup: {adaptor: {innerHTML(node: MathJaxNode): string}};
  };
  export default MathJax;
}
