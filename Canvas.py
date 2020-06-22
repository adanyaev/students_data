from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class Graphic(FigureCanvas):
    def __init__(self, figure):
        self.fig = figure
        FigureCanvas.__init__(self, self.fig)
