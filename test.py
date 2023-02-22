import networkx as nx

G=nx.DiGraph()
G.add_node('q')
G.add_node(1)
G.add_node(2)
G.add_edge('q',3)


from matplotlib import pyplot as plt
import networkx as nx
nx.draw_networkx(G)
plt.show()
