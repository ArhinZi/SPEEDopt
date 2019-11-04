from base_optimizer import BaseOptimizer
from benderopt.base import OptimizationProblem, Observation
import benderopt
from benderopt.optimizer.parzen_estimator import ParzenEstimator

# From https://github.com/Dreem-Organization/benderopt


class BenderOptimizer(BaseOptimizer):

    def box(self, **kwargs):
        context = [kwargs[key] for key in kwargs.keys()]
        return super().box(context)

    def run(self, limits, minimize, dynamic_optimizing, field_for_optimizing, project_path, scope, batches):
        super().run(minimize, dynamic_optimizing, field_for_optimizing, project_path)

        optimization_problem_parameters = []
        init = []
        bounds = []
        for key in limits.keys():
            pars = {
                "name": key,
                "category": "uniform",
                "search_space": {
                    "low": limits[key]["low"],
                    "high": limits[key]["high"],
                    "step": limits[key]["step"]
                }
            }
            self.fields.append(key)
            optimization_problem_parameters.append(pars)

        best_sample = benderopt.minimize(
            f=self.box,
            optimization_problem_parameters=optimization_problem_parameters,
            number_of_evaluation=scope,
            debug=False
        )
        return best_sample