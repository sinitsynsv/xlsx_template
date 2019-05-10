class Context:
    def __init__(self, context_data, env):
        self.context_data = context_data
        self.env = env

    def resolve(self, name):
        if name in self.context_data:
            return self.env.resolve(self.context_data[name], name, True)
        else:
            return self.env.resolve(None, name, False)
