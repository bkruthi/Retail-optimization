declare module "javascript-lp-solver" {
  const Solver: {
    Solve: (model: unknown) => Record<string, unknown>;
  };

  export default Solver;
}