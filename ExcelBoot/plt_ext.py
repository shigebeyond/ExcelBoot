import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np
import os
import uuid

# matplotlib支持中文显示
def fix_cn():
    if 'SimHei' in plt.rcParams['font.sans-serif']:
        return

    # 1 添加字体
    # 先查找黑体字体, 如果找不到, 默认返回DejaVuSans.ttf
    font_path = font_manager.fontManager.findfont("SimHei")
    # 如果找不到,则添加字体
    if font_path.endswith('DejaVuSans.ttf'):
        ttf_file = __file__.replace("plt_ext.py", "simhei.ttf")
        font_manager.fontManager.addfont(ttf_file)
    # 输出字体
    # for f in font_manager.fontManager.ttflist:
    #     print(f.name)

    # 2 应用字体
    # plt.rcParams['font.family'] = ['sans-serif']  # 可省
    plt.rcParams['font.sans-serif'] = ['SimHei'] + plt.rcParams['font.sans-serif']  # 设置中文字体
    plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

# 绘图
def plot(df, **kwargs):
    """
    Make plots of Series or DataFrame.

    Uses the backend specified by the
    option ``plotting.backend``. By default, matplotlib is used.

    Parameters
    ----------
    df : Series or DataFrame
        The object for which the method is called.
    x : label or position, default None
        Only used if data is a DataFrame.
    y : label, position or list of label, positions, default None
        Allows plotting of one column versus another. Only used if data is a
        DataFrame.
    kind : str
        The kind of plot to produce:

        - 'line' : line plot (default)
        - 'bar' : vertical bar plot
        - 'barh' : horizontal bar plot
        - 'hist' : histogram
        - 'box' : boxplot
        - 'kde' : Kernel Density Estimation plot
        - 'density' : same as 'kde'
        - 'area' : area plot
        - 'pie' : pie plot
        - 'scatter' : scatter plot
        - 'hexbin' : hexbin plot.
    ax : matplotlib axes object, default None
        An axes of the current figure.
    subplots : bool, default False
        Make separate subplots for each column.
    sharex : bool, default True if ax is None else False
        In case ``subplots=True``, share x axis and set some x axis labels
        to invisible; defaults to True if ax is None otherwise False if
        an ax is passed in; Be aware, that passing in both an ax and
        ``sharex=True`` will alter all x axis labels for all axis in a figure.
    sharey : bool, default False
        In case ``subplots=True``, share y axis and set some y axis labels to invisible.
    layout : tuple, optional
        (rows, columns) for the layout of subplots.
    figsize : a tuple (width, height) in inches
        Size of a figure object.
    use_index : bool, default True
        Use index as ticks for x axis.
    title : str or list
        Title to use for the plot. If a string is passed, print the string
        at the top of the figure. If a list is passed and `subplots` is
        True, print each item in the list above the corresponding subplot.
    grid : bool, default None (matlab style default)
        Axis grid lines.
    legend : bool or {'reverse'}
        Place legend on axis subplots.
    style : list or dict
        The matplotlib line style per column.
    logx : bool or 'sym', default False
        Use log scaling or symlog scaling on x axis.
        .. versionchanged:: 0.25.0

    logy : bool or 'sym' default False
        Use log scaling or symlog scaling on y axis.
        .. versionchanged:: 0.25.0

    loglog : bool or 'sym', default False
        Use log scaling or symlog scaling on both x and y axes.
        .. versionchanged:: 0.25.0

    xticks : sequence
        Values to use for the xticks.
    yticks : sequence
        Values to use for the yticks.
    xlim : 2-tuple/list
        Set the x limits of the current axes.
    ylim : 2-tuple/list
        Set the y limits of the current axes.
    xlabel : label, optional
        Name to use for the xlabel on x-axis. Default uses index name as xlabel.

        .. versionadded:: 1.1.0

    ylabel : label, optional
        Name to use for the ylabel on y-axis. Default will show no ylabel.

        .. versionadded:: 1.1.0

    rot : int, default None
        Rotation for ticks (xticks for vertical, yticks for horizontal
        plots).
    fontsize : int, default None
        Font size for xticks and yticks.
    colormap : str or matplotlib colormap object, default None
        Colormap to select colors from. If string, load colormap with that
        name from matplotlib.
    colorbar : bool, optional
        If True, plot colorbar (only relevant for 'scatter' and 'hexbin'
        plots).
    position : float
        Specify relative alignments for bar plot layout.
        From 0 (left/bottom-end) to 1 (right/top-end). Default is 0.5
        (center).
    table : bool, Series or DataFrame, default False
        If True, draw a table using the data in the DataFrame and the data
        will be transposed to meet matplotlib's default layout.
        If a Series or DataFrame is passed, use passed data to draw a
        table.
    yerr : DataFrame, Series, array-like, dict and str
        See :ref:`Plotting with Error Bars <visualization.errorbars>` for
        detail.
    xerr : DataFrame, Series, array-like, dict and str
        Equivalent to yerr.
    stacked : bool, default False in line and bar plots, and True in area plot
        If True, create stacked plot.
    sort_columns : bool, default False
        Sort column names to determine plot ordering.
    secondary_y : bool or sequence, default False
        Whether to plot on the secondary y-axis if a list/tuple, which
        columns to plot on secondary y-axis.
    mark_right : bool, default True
        When using a secondary_y axis, automatically mark the column
        labels with "(right)" in the legend.
    include_bool : bool, default is False
        If True, boolean values can be plotted.
    backend : str, default None
        Backend to use instead of the backend specified in the option
        ``plotting.backend``. For instance, 'matplotlib'. Alternatively, to
        specify the ``plotting.backend`` for the whole session, set
        ``pd.options.plotting.backend``.

        .. versionadded:: 1.0.0

    **kwargs
        Options to pass to matplotlib plotting method.

    Returns
    -------
    :class:`matplotlib.axes.Axes` or numpy.ndarray of them
        If the backend is not the default matplotlib one, the return value
        will be the object returned by the backend.

    Notes
    -----
    - See matplotlib documentation online for more on this subject
    - If `kind` = 'bar' or 'barh', you can specify relative alignments
      for bar plot layout by `position` keyword.
      From 0 (left/bottom-end) to 1 (right/top-end). Default is 0.5
      (center)
    """
    fix_cn()

    # 1 准备参数
    # 1.1 指定y轴用到的几个列
    if 'y' in kwargs:
        kwargs['y'] = kwargs['y'].split(',')
    # 1.2 不同图的默认参数
    default_args = {}
    kind = kwargs['kind'] # 图类型
    if kind == 'hist': # 直方图: bins柱个数,bins越大表示每个柱的宽度越小
        default_args = {'subplots': True, 'bins': 20, 'sharex': False, 'layout': (1, len(df.columns))}
    elif kind == 'kde': # 核密度图
        default_args = {'grid': True}
    elif kind == 'pie': # 饼图: figsize(长度,宽度), startangle起始角度,  radius饼图的半径, legend图例说明(颜色+中文说明)
        default_args = {'subplots': True, 'figsize': (12, 8), 'autopct': '%.2f%%', 'radius': 1.2,
                 'startangle': 250, 'legend': False, 'colormap': 'viridis'}
        '''
        fix bug: TypeError: '<' not supported between instances of 'str' and 'int'
        原因: PiePlot构造函数中会检查 `if (data < 0).any().any():`, 会对df中所有列数据进行>0操作, 而对于x列一般是str类型, 执行>0操作时报错
        解决: 直接干掉x列, 将他弄成index
        '''
        if 'x' in kwargs:
            x_col = kwargs['x']
            del kwargs['x']
            # x列作为index
            df2 = df.set_index(df[x_col])
            del df2[x_col]
            df = df2
    # 合并参数
    kwargs = dict(default_args, **kwargs)

    # 2 绘图
    plot = df.plot(**kwargs)
    # 可能是多层的plot的ndarray, 要获得最终的plot用于保存图片
    while isinstance(plot, np.ndarray):
        plot = plot[0]

    # 3 保存成图片
    # https://blog.csdn.net/Hodors/article/details/109199595
    fig = plot.get_figure()
    path = prepare_img_path(kind)
    fig.savefig(path)
    return path

# 准备图片路径
def prepare_img_path(kind):
    id = str(uuid.uuid1())
    dir = os.path.abspath('tmp')
    if not os.path.exists(dir):
        os.makedirs(dir)
    path = dir + f"/{kind}-{id}.png"
    return path
