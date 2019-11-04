from base_optimizer import BaseOptimizer
import blackbox as bb


class BlackBoxOptimizer(BaseOptimizer):

    def run(self, limits, minimize, dynamic_optimizing, field_for_optimizing, project_path, scope, batches):
        super().run(minimize, dynamic_optimizing, field_for_optimizing, project_path)

        domain = []
        for key in limits.keys():
            self.fields.append(key)
            domain.append(limits[key])

        bb.search_min(
            f=self.box,  # given function
            domain=domain,  # ranges of each parameter
            budget=scope,  # total number of function calls available
            batch=batches,  # number of calls that will be evaluated in parallel
            resfile='output.csv'  # text file where results will be saved
        )
