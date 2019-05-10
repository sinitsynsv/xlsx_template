class LoopContext:
    def __init__(self, items):
        if not hasattr(items, "__len__"):
            items = list(items)
        self.items = items
        self.index0 = -1
        self.length = len(items)

    def __iter__(self):
        for item in self.items:
            self.index0 += 1
            yield item

    @property
    def index(self):
        return self.index0 + 1

    @property
    def revindex(self):
        return self.length - self.index0

    @property
    def revindex0(self):
        return self.length - self.index

    @property
    def first(self):
        return self.index0 == 0

    @property
    def last(self):
        return self.index0 == self.length - 1
