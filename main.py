from LoadFlow import *
import pandas as pd
import time

# make a pandas based function to generate the nodes and line objects

def main():
    # n1 = Node(1, 1.05, 0, 0.0, 0, 0, 0, 0, 0)
    # n2 = Node(2, 1.05, 0, 0.5, 0, 0, 0, -0.5, 1.0)
    # n3 = Node(2, 1.07, 0, 0.6, 0, 0, 0, -0.5, 1.5)
    # n4 = Node(3, 1.0, 0, 0.0, 0, 0.7, 0.7, 0, 0)
    # n5 = Node(3, 1.0, 0, 0.0, 0, 0.7, 0.7, 0, 0)
    # n6 = Node(3, 1.0, 0, 0.0, 0, 0.7, 0.7, 0, 0)

    # l1 = Line(n1, n2, 0.1, 0.2, 0.02, 1)
    # l2 = Line(n1, n4, 0.05, 0.2, 0.02, 1)
    # l3 = Line(n1, n5, 0.08, 0.3, 0.03, 1)
    # l4 = Line(n2, n3, 0.05, 0.25, 0.03, 1)
    # l5 = Line(n2, n4, 0.05, 0.1, 0.01, 1)
    # l6 = Line(n2, n5, 0.1, 0.3, 0.02, 1)
    # l7 = Line(n2, n6, 0.07, 0.2, 0.025, 1)
    # l8 = Line(n3, n5, 0.12, 0.26, 0.025, 1)
    # l9 = Line(n3, n6, 0.02, 0.1, 0.01, 1)
    # l10 = Line(n4, n5, 0.2, 0.4, 0.04, 1)
    # l11 = Line(n5, n6, 0.1, 0.3, 0.03, 1)

    # nodes = [n1, n2, n3, n4, n5, n6]
    # lines = [l1, l2, l3, l4, l5, l6, l7, l8, l9, l10, l11]

    # grid = Grid(nodes, lines)
    # grid.loadflow()
    # grid.printResults()

    file = 'data/case57.xlsx'
    busData = pd.read_excel(file, sheet_name='busData')
    lineData = pd.read_excel(file, sheet_name='branchData')
    bnum = len(busData.axes[0])
    lnum = len(lineData.axes[0])

    nodes = []
    lines = []

    for i in range(bnum):
        nodes.append(
            Node(
                busData.at[i, 'type'], busData.at[i, 'Vm'], busData.at[i, 'Va'], busData.at[i, 'Pg'], 
                busData.at[i, 'Qg'], busData.at[i, 'Pd'], busData.at[i, 'Qd'], busData.at[i, 'Qmin'], 
                busData.at[i, 'Qmax']
            )
        )
    
    for i in range(lnum):
        lines.append(
            Line(
                nodes[lineData.at[i, 'fbus']-1], nodes[lineData.at[i, 'tbus']-1], 
                lineData.at[i, 'r'], lineData.at[i, 'x'], lineData.at[i, 'b']
            )
        )


    grid = Grid(nodes=nodes, lines=lines)
    start = time.time()
    grid.NR()
    end = time.time()
    grid.printResults()
    print("\nTime taken: %f" % (end-start))

    # print(bnum)
    # print(lnum)
    # print(lineData.at[0,'tbus'])



if __name__ == '__main__':
    main()