from blackbox_optimizer import BlackBoxOptimizer
from scipy_optimizer import SciPyOptimizer
from brain_optimizer import BrainBBOptimizer
from bender_optimizer import BenderOptimizer

if __name__ == "__main__":
    # opt = SciPyOptimizer()
    # opt = BrainBBOptimizer()
    # opt = BenderOptimizer()
    opt = BlackBoxOptimizer()
    res = opt.run(
        project_path="E:\\SPEED\\Projects\\test2.bd4",
        dynamic_optimizing=True,
        minimize=False,
        scope=100,
        batches=4,
        field_for_optimizing='Tshaft',
        limits={
            "BetaM": {
                "low": 100,
                "high": 180,
                "step": 1
            },
            "Gap": {
                "low": 1,
                "high": 10,
                "step": 0.1
            },
            "LM": {
                "low": 1,
                "high": 10,
                "step": 0.1
            },
        }

    )
    print(res)



