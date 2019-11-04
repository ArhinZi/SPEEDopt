from blackbox_optimizer import BlackBoxOptimizer
from scipy_optimizer import SciPyOptimizer

if __name__ == "__main__":
    opt = BlackBoxOptimizer()
    opt.run(
        project_path="E:\\SPEED\\Projects\\test2.bd4",
        dynamic_optimizing=True,
        minimize=False,
        scope=100,
        batches=4,
        field_for_optimizing='Tshaft',
        limits={
            "Gap": [1, 2],
            "LM": [1, 10],
            "BetaM": [160, 180]
        }

    )



