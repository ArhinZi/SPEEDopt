from speed import speedWrapper
import blackbox as bb


class BaseOptimizer:

    def __init__(self):
        self.minimize = None
        self.dynamic_optimizing = None
        self.field_for_optimizing = None
        self.pcbdc_instance = None
        self.project_path = None
        self.fields = []

    def get_instance(self):
        self.pcbdc_instance = speedWrapper(self.project_path)

    def box(self, context):
        self.get_instance()
        params = {field: context[i] for (i, field) in enumerate(self.fields)}
        self.pcbdc_instance.setParams(params)

        if self.dynamic_optimizing:
            self.pcbdc_instance.Dynamic()
        else:
            self.pcbdc_instance.Static()

        target = self.pcbdc_instance.getParam(self.field_for_optimizing)

        print(context, target)
        if not self.minimize:
            # if target <= 0:
            #     target = 999
            # else:
            #     target = 1/target
            target = -1*target

        return target

    def run(self, minimize, dynamic_optimizing, field_for_optimizing, project_path):
        self.minimize = minimize
        self.dynamic_optimizing = dynamic_optimizing
        self.field_for_optimizing = field_for_optimizing
        self.project_path = project_path

    def exit(self):
        self.pcbdc_instance.Quit()
