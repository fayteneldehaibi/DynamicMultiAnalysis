import sys
import os

import graphviz
import pandas
from pandas import ExcelWriter
import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
import matplotlib
from matplotlib import pyplot as plt
from matplotlib.figure import Figure
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import statistics
import math
import scipy.stats
from scipy.stats import ttest_ind
import numpy as np
import re
import hypernetx
import warnings

# AUTHORS: ASHTI M. SHAH, FAYTEN EL-DEHAIBI
# MENTORS: DR. YORAM VODOVOTZ AND DR. RUBEN ZAMORA
# DATE: January, 2023
# GUI Added Apr 2023 by Fayten El-Dehaibi

app = QtWidgets.QApplication(sys.argv)

class Form(QtWidgets.QDialog):
    #GUI Initialization
    def __init__(self, parent = None):
        super(Form, self).__init__(parent)

        self.fileLine = QtWidgets.QLineEdit()
        self.fileLine.setReadOnly(True)
        self.fileButton = QtWidgets.QPushButton('...')
        self.sheetLine = QtWidgets.QLineEdit('Sheet1')
        dynaAndDyHyp = QtWidgets.QRadioButton('Analyze DyNA and DyHyp for each tissue')
        dynaAndDyHyp.setChecked(True)
        dyHypNetwork = QtWidgets.QRadioButton('Analyze DyHyp Network for all tissues')
        self.groupColumn = QtWidgets.QLineEdit('Condition')
        self.baselineGroupName = QtWidgets.QLineEdit('Baseline')
        targetGroupName = QtWidgets.QLineEdit()
        self.timeColumn = QtWidgets.QLineEdit('Time')
        tissueColumn = QtWidgets.QLineEdit('Compartment')
        dynaThreshLine = QtWidgets.QDoubleSpinBox()
        dynaThreshLine.setValue(0.90)
        dynaThreshLine.setMinimum(0.0)
        dynaThreshLine.setMaximum(1.0)
        dyhypStDevLine = QtWidgets.QSpinBox()
        dyhypStDevLine.setValue(1)
        dyhypStDevLine.setMinimum(0)
        outputTitle = QtWidgets.QLineEdit()
        self.runButton = QtWidgets.QPushButton('Analyze DyNA and DyHyp')

        #Setting Gui Layout
        fileLayout = QtWidgets.QHBoxLayout()
        fileLayout.addWidget(QtWidgets.QLabel('File:'))
        fileLayout.addWidget(self.fileLine)
        fileLayout.addWidget(self.fileButton)
        pageLayout = QtWidgets.QHBoxLayout()
        pageLayout.addWidget(QtWidgets.QLabel('Sheet:'))
        pageLayout.addWidget(self.sheetLine)
        groupLO = QtWidgets.QHBoxLayout()
        groupLO.addWidget(QtWidgets.QLabel('Group Column Name:'))
        groupLO.addWidget(self.groupColumn)
        baseLO = QtWidgets.QHBoxLayout()
        baseLO.addWidget(QtWidgets.QLabel('Control Group Name:'))
        baseLO.addWidget(self.baselineGroupName)
        targetLO = QtWidgets.QHBoxLayout()
        targetLO.addWidget(QtWidgets.QLabel('Target Group Name:'))
        targetLO.addWidget(targetGroupName)
        tissueLO = QtWidgets.QHBoxLayout()
        tissueLO.addWidget(QtWidgets.QLabel('Tissue Column Name:'))
        tissueLO.addWidget(tissueColumn)
        timeLO = QtWidgets.QHBoxLayout()
        timeLO.addWidget(QtWidgets.QLabel('Time Column Name:'))
        timeLO.addWidget(self.timeColumn)
        threshLO = QtWidgets.QHBoxLayout()
        threshLO.addWidget(QtWidgets.QLabel('DyNA Threshold:'))
        threshLO.addWidget(dynaThreshLine)
        stdevLO = QtWidgets.QHBoxLayout()
        stdevLO.addWidget(QtWidgets.QLabel('DyHyp Standard Deviations:'))
        stdevLO.addWidget(dyhypStDevLine)
        outputLO = QtWidgets.QHBoxLayout()
        outputLO.addWidget(QtWidgets.QLabel('Output File Prefix:'))
        outputLO.addWidget(outputTitle)
        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(fileLayout)
        layout.addLayout(pageLayout)
        layout.addWidget(QtWidgets.QLabel(''))
        layout.addWidget(QtWidgets.QLabel('Analysis Method'))
        layout.addWidget(dynaAndDyHyp)
        layout.addWidget(dyHypNetwork)
        layout.addWidget(QtWidgets.QLabel(''))
        layout.addWidget(QtWidgets.QLabel('Enter Column Names and Groups to Analyze'))
        layout.addLayout(groupLO)
        layout.addLayout(baseLO)
        layout.addLayout(targetLO)
        layout.addLayout(tissueLO)
        layout.addLayout(timeLO)
        layout.addWidget(QtWidgets.QLabel(''))
        layout.addWidget(QtWidgets.QLabel('Analysis Parameters'))
        layout.addLayout(threshLO)
        layout.addLayout(stdevLO)
        layout.addWidget(QtWidgets.QLabel(''))
        layout.addLayout(outputLO)
        layout.addWidget(self.runButton)
        self.setLayout(layout)
        self.setWindowTitle('DyNA + DyHypergraphs')

        #Signals and Functions
        self.fileButton.clicked.connect(lambda: self.getPath(self.fileLine))
        self.runButton.clicked.connect(lambda: self.DyHyp_Network_Complexity(dyhypStDevLine.value(), targetGroupName.text(),
                                                                             tissueColumn.text(), outputTitle.text()) if dyHypNetwork.isChecked()
                                               else self.runDyNAandDyHyp(dynaThreshLine.value(), dyhypStDevLine.value(),
                                                                         targetGroupName.text(), tissueColumn.text(), outputTitle.text()))

    #Pulls up file window to select input data file
    def getPath(self,path):
        fileBox = QtWidgets.QFileDialog(self)
        inFile = QtCore.QFileInfo(fileBox.getOpenFileName(None,filter='*.xls* *.csv')[0])
        filepath = inFile.absoluteFilePath()
        if any([filepath.endswith('.xls'),filepath.endswith('.xlsx'),filepath.endswith('.csv')]):
            path.setText(filepath)

    #Reads input data file from filepath PATH and sheetname SHEET, then applies filters for group conditions
    def readFile(self,targetGroup):
        path = self.fileLine.text()
        sheet = self.sheetLine.text()
        columnName = self.groupColumn.text()
        baselineName = self.baselineGroupName.text()
        raw_data = pandas.DataFrame()
        if(path.endswith('.xls') or path.endswith('.xlsx')):
            raw_data = pandas.read_excel(path,sheet_name=sheet)
        if(path.endswith('.csv')):
            df = pandas.read_csv(path,engine='python',header=0,iterator=True,
                                 chunksize=15000,infer_datetime_format=True)
            raw_data = pandas.DataFrame(pandas.concat(df,ignore_index=True))
        # "Organize data by condition"
        raw_data = raw_data.dropna(axis='columns',how='all')
        all_data_Baseline = raw_data.loc[raw_data[columnName] == baselineName]
        all_data = pandas.concat([all_data_Baseline, raw_data.loc[raw_data[columnName] == targetGroup]])
        return all_data

    #Checks for missing user variables
    def errorChecks(self,condition, tissue_column):
        if any([self.fileLine.text().isspace(),
                self.groupColumn.text().isspace(),
                self.baselineGroupName.text().isspace(),
                self.timeColumn.text().isspace(),
                condition.isspace(),
                tissue_column.isspace()]):
            errMessage = QtWidgets.QMessageBox.warning(self,'Error: Blank fields detected',
                                            'One or more fields are empty. Please enter all information and try again.',
                                            '',defaultButtonNumber=0)
            return True
        else:
            return False

    def sortTime(self,df):
        baseList = list(set(df.loc[:,self.timeColumn.text()]))
        numList = []
        day_mod = re.compile(r'd',re.IGNORECASE)
        hour_mod = re.compile(r'[hr]',re.IGNORECASE)
        for b in baseList:
            if 'd' in b:
                numList.append(float(day_mod.sub('',b)))
            elif 'h' in b:
                numList.append(float(hour_mod.sub('',b))/24)
        intList = sorted(numList)
        strList = [str(s) + 'd' for s in intList]
        return strList,intList

    def findFirstCyt(self,df,tissue_col):
        colFloatAll = list(df.select_dtypes(include='number').columns)
        firstCytCol = [col for col in colFloatAll if col not in [self.groupColumn.text(),self.timeColumn.text(),tissue_col]][0]
        return df.columns.get_loc(firstCytCol)

    # Function to get the correlation of each inflammatory mediator in a single organ
    # with TIME based on RATE OF CHANGE
    def get_rate_of_change_inflammatory_mediator_and_TIME(self,tissue_rel_data, time_int,startCyt):
        # tissue_rel_data: rel_data_cur_organ; table of the relevant data from two time points from a single organ
        # time_int: int array of consective time points. ex [1, 3]
        # startCyt: int loc of first column of mediators

        #"Get the data frames of data for each time point"
        time_1 = tissue_rel_data.loc[tissue_rel_data[self.groupColumn.text()] == self.baselineGroupName.text()]
        time_2 = tissue_rel_data.loc[tissue_rel_data[self.groupColumn.text()] != self.baselineGroupName.text()]
        #"Get the data frames of just the inflammatory mediators for each time point"

        time1_mediators = time_1.iloc[:, startCyt:]
        time2_mediators = time_2.iloc[:, startCyt:]

        cytokines = list(time1_mediators.columns.values)  # list of all cytokines

        #"Calculate rate of change"
        rate_of_change_table = pandas.DataFrame()
        for c in cytokines:
            x1 = time_int[0]
            x2 = time_int[1]
            #"Rate of change is calculated using median cytokine value at each timepoint"
            y1 = time1_mediators[c].median()
            y2 = time2_mediators[c].median()
            #"Calculate rate of change, x = time, y = median inflammatory value"
            rate_of_change = (y2 - y1) / (x2 - x1)
            rate_of_change_table[c] = [rate_of_change]
        return (rate_of_change_table)

    # Function to return a list of mediators that are positively or negatively correlated with time in a single organ
    def get_significant_mediators_withTIME(self,inflammatory_mediator_rates, num_std_dev):
        # inflammatory_mediator_rates: correlation matrix (1x20) which has the pearson correlation of each inflammatory
        # mediator with itself over a dynamic time interval
        # num_std_dev: number of standard deviations above the mean rate of change at which a cytokine is considered
        # to be significantly increasing

        cytokines = list(inflammatory_mediator_rates.columns.values)  # list of all cytokines
        pos_mediators = []
        neg_mediators = []

        for c in cytokines:
            if inflammatory_mediator_rates[c][0] > 0:
                pos_mediators.append(inflammatory_mediator_rates[c][0])
            elif inflammatory_mediator_rates[c][0] < 0:
                neg_mediators.append(inflammatory_mediator_rates[c][0])

        warnings.filterwarnings('ignore', category=RuntimeWarning, module='numpy')#silences divide by zero warnings
        meanPos = np.mean(pos_mediators)
        stdevPos = np.std(pos_mediators)
        thresholdPos = meanPos + stdevPos * num_std_dev
        meanNeg = np.mean(neg_mediators)
        stdevNeg = np.std(neg_mediators)
        thresholdNeg = meanNeg - stdevNeg * num_std_dev

        significant_pos_mediators = []
        significant_neg_mediators = []
        for j in cytokines:
            if inflammatory_mediator_rates[j][0] > thresholdPos:
                significant_pos_mediators.append(j)
            if inflammatory_mediator_rates[j][0] < thresholdNeg:
                significant_neg_mediators.append(j)
        return significant_pos_mediators, significant_neg_mediators

    # Function to get the corelation every pair of inflammatory mediators in a single organ across two time points
    def get_correlation_matrix_withMEDIATORS_individual_tissue(self,tissue_rel_data, time_int,startCyt):
        # baseline_data: data at t0
        # tissue_rel_data: rel_data_cur_organ; table of the relevant data from two time points from a single organ
        # time_int: int array of consective time points. ex [1, 3]
        # startCyt: int loc of first column of mediators
        #"Get the data frames of data for each time point"
        time_1 = tissue_rel_data.loc[tissue_rel_data[self.groupColumn.text()] == self.baselineGroupName.text()]
        time_2 = tissue_rel_data.loc[tissue_rel_data[self.groupColumn.text()] != self.baselineGroupName.text()]

        #"Get the data frames of just the inflammatory mediators for each time point"
        time1_mediators = time_1.iloc[:, startCyt:]
        time2_mediators = time_2.iloc[:, startCyt:]

        cytokines = list(time1_mediators.columns.values)  # list of all cytokines

        corr_table_all_mediators = pandas.DataFrame(columns=range(len(cytokines)))
        corr_table_all_mediators.columns = cytokines
        #"Nested loop in which we correlate x cytokine with all other cytokines"
        for x in cytokines:
            #"x: the cytokine that we are correlating all other cytokines with"
            corr_data_x_current_mediator = pandas.concat([time1_mediators[x],time2_mediators[x]],ignore_index=True)
            #corr_data_x_current_mediator = pandas.DataFrame(time1_mediators[x].append(time2_mediators[x])).reset_index(drop=True)
            correl_with_cur_mediator = pandas.DataFrame()
            for c in cytokines:
                #corr_data_y = pandas.DataFrame(time1_mediators[c].append(time2_mediators[c])).reset_index(drop=True)
                corr_data_y = pandas.concat([time1_mediators[c],time2_mediators[c]],ignore_index=True)
                correlation_data = pandas.concat([corr_data_x_current_mediator, corr_data_y], axis=1, ignore_index=True)
                correlation_data.columns = [x, c]
                #"to allow a DyNA edge to be drawn, both cytokine X and C must be significantly different than the baseline value"
                #"by default, assume that X and C are not correlated"
                correlation_coeff_cur_cytokine = 0
                pearsons_correl_cur_cytokine = correlation_data.corr(method='pearson')
                #"Check if X is significantly different from baseline"
                t_test_time2_x,pvalue_x = ttest_ind(time1_mediators[x], time2_mediators[x], nan_policy='omit')
                t_test_time2_c,pvalue_c = ttest_ind(time1_mediators[c], time2_mediators[c], nan_policy='omit')
                if (pvalue_x < 0.05) and (pvalue_c < 0.05):
                    correlation_coeff_cur_cytokine = pearsons_correl_cur_cytokine.iloc[0, 1]
                correl_with_cur_mediator[c] = [correlation_coeff_cur_cytokine]
            corr_table_all_mediators = pandas.concat([corr_table_all_mediators, correl_with_cur_mediator])
        corr_table_all_mediators.index = cytokines
        return (corr_table_all_mediators)

    # Function return a dictionary of significant mediators that are positively or negatively correlated with each other
    def get_significant_mediators_withEACHOTHER(self, correl_matrix_with_mediators, threshold):
        # correl_matrix_with_mediators: correlation matrix (27x27) which has the pearson correlation of each inflammatory
        # mediator with every other mediator over a dynamic time interval
        # threshold: minimum pearson correlation, type: float
        cytokines_column = list(correl_matrix_with_mediators.columns.values)
        cytokines_row = list(correl_matrix_with_mediators.index.values)  # list of all cytokines
        significant_mediators_pos = {}
        significant_mediators_neg = {}
        for x in cytokines_column:
            significant_mediators_in_col_p = []
            significant_mediators_in_col_n = []
            for c in cytokines_row:
                cur_correlation = correl_matrix_with_mediators.loc[c, x]
                if math.isnan(cur_correlation) == False:
                    if cur_correlation >= threshold and x != c:
                        significant_mediators_in_col_p.append(c)
                    elif cur_correlation <= threshold * -1 and x != c:
                        significant_mediators_in_col_n.append(c)
            if len(significant_mediators_in_col_p) > 0:
                significant_mediators_pos[x] = significant_mediators_in_col_p
            if len(significant_mediators_in_col_n) > 0:
                significant_mediators_neg[x] = significant_mediators_in_col_n
        return significant_mediators_pos, significant_mediators_neg

    def group_edges_dyHyp(self,tissue_1, tissue_2, tissue_1_name, tissue_2_name):
        # Tissue_1: significant mediators within plasma
        # Tissue_2: significant mediators within another tissue
        # tissue_1_name: str, tissue1
        # tissue_2_name: str, tissue2
        edges = {tissue_1_name: [], tissue_2_name: [], "{} and {}".format(tissue_1_name, tissue_2_name): []}
        for j in tissue_1:
            if j in tissue_2:
                edges["{} and {}".format(tissue_1_name, tissue_2_name)].append(j)
                tissue_2.remove(j)
            else:
                edges[tissue_1_name].append(j)
        if len(tissue_2) != 0:
            for n in tissue_2:
                edges[tissue_2_name].append(n)
        return edges

    # Function to calculate DyNA Network Complexity
    def get_dyNA_network_complexity(self,dict_mediators_connections_pos, dict_mediators_connections_neg):
        sum_connections_TOTAL = 0
        if len(dict_mediators_connections_pos.keys()) != 0:
            for mediator in (dict_mediators_connections_pos.keys()):
                sum_connections_TOTAL += len(dict_mediators_connections_pos[mediator])
        if len(dict_mediators_connections_neg.keys()) != 0:
            for mediator in (dict_mediators_connections_neg.keys()):
                sum_connections_TOTAL += len(dict_mediators_connections_neg[mediator])
        network_complexity = (sum_connections_TOTAL / 2) / 200
        return network_complexity

    #Draws the mediator network within a single tissue
    def draw_intratissue_network(self,title,sigMediators):
        #tissue: str, name of tissue + Pos/Neg + time point
        #sigMediators: dict of significantly correlated mediators within given tissue
        if len(sigMediators) == 0:
            return 0
        mediators = list(sigMediators.keys())
        network = graphviz.Digraph('network',strict=True)
        network.attr(rank='same')
        for m in mediators:
            network.node(''.join(m.split(' ')))
        allEdges = []
        for k in sigMediators.keys():
            if len(sigMediators[k]) > 0:
                for val in sigMediators[k]:
                    if ''.join((str(val)+str(k)).split(' ')) not in allEdges:
                        network.edge(''.join(str(k).split(' ')),''.join(str(val).split(' ')),dir='both')
                        allEdges.append(''.join((str(k)+str(val)).split(' ')))
        result_ = network.render(title+'.png',cleanup=True,format='png',engine='dot').replace('\\', '/')
        return 1

    def draw_intertissue_network(self,subnets):
        return 0
        # make boxy hypergraph of tissues using matplotlib Polycollections....uh...that may take some doing
        #ax = plt.subplot()
        #read intratissue networks previously generated
        #pic = plt.imread('<tissue>.png')
        #embed subnets into proper spots
        #offImage = OffsetImage(pic,zoom=0.45)#work on zoom later
        #tissueBox = AnnotationBbox(offImage, (<coordinates>),xycoords='axes fraction', box_alignment=(1.1,-0.1)) #haha, get it? ... you get it?
        #ax.addartist(tissue_Box)
        #plt.savefig()

    # Function to create graph of DyHyp an DyNA
    def visual_graph(self,figure, grid_spec, grouped_edges_dyHyp, sigMediators_tissue1, sigMediators_tissue2, tissue_1_name,
                     tissue_2_name):
        # grouped_edges_dyHyp: dictionary of edges sorted by tissues and combined tissues
        # sigMediators_tissue1: dictionary of mediators significantly correlated with each other in tissue1
        # sigMediators_tissue2: dictionary of mediators significantly correlated with each other in tissue2
        # tissue_1_name: str, tissue1
        # tissue_2_name: str, tissue2
        # title: title of plot

        # display(grouped_edges_dyHyp)
        ax = figure.add_subplot(grid_spec[0, 0])
        x_coords_node = [0.75, 6]
        tissue_1_ycoord = [5, 5]
        tissue_2_ycoord = [0.1, 0.1]
        tissues_1_and_2_ycoord1 = [5, 2.5]
        tissues_1_and_2_ycoord2 = [0.1, 2.5]
        plt.text(-2, 5.2, tissue_1_name, fontsize=36, verticalalignment='top')
        plt.text(-2, 0.2, tissue_2_name, fontsize=36, verticalalignment='top')
        plt.plot(x_coords_node, tissue_1_ycoord, 'k-', 8)
        plt.plot(x_coords_node, tissue_2_ycoord, 'k-', 8)
        if (grouped_edges_dyHyp[tissue_1_name]) != []:
            if len(grouped_edges_dyHyp[tissue_1_name]) > 5:
                plt.text(6.5, 5.3, "           ".join(grouped_edges_dyHyp[tissue_1_name][0:6]), fontsize=24,
                         verticalalignment='top')
                plt.text(6.5, 5, "           ".join(grouped_edges_dyHyp[tissue_1_name][6:]), fontsize=24,
                         verticalalignment='top')
            else:
                plt.text(6.5, 5, "           ".join(grouped_edges_dyHyp[tissue_1_name]), fontsize=24,
                         verticalalignment='top')
        if (grouped_edges_dyHyp[tissue_2_name]) != []:
            if len(grouped_edges_dyHyp[tissue_1_name]) > 5:
                plt.text(6.5, 0.4, "           ".join(grouped_edges_dyHyp[tissue_2_name][0:6]), fontsize=24,
                         verticalalignment='top')
                plt.text(6.5, 0.1, "           ".join(grouped_edges_dyHyp[tissue_2_name][6:]), fontsize=24,
                         verticalalignment='top')
            else:
                plt.text(6.5, 0.1, "           ".join(grouped_edges_dyHyp[tissue_2_name]), fontsize=24,
                         verticalalignment='top')
        if (grouped_edges_dyHyp["{} and {}".format(tissue_1_name, tissue_2_name)]) != []:
            plt.plot(x_coords_node, tissues_1_and_2_ycoord1, 'k-', 8)
            plt.plot(x_coords_node, tissues_1_and_2_ycoord2, 'k-', 8)
            if len(grouped_edges_dyHyp[tissue_1_name]) > 5:
                plt.text(6.5, 2.8,
                         "           ".join(grouped_edges_dyHyp["{} and {}".format(tissue_1_name, tissue_2_name)][0:6]),
                         fontsize=24, verticalalignment='top')
                plt.text(6.5, 2.5,
                         "           ".join(grouped_edges_dyHyp["{} and {}".format(tissue_1_name, tissue_2_name)][6:]),
                         fontsize=24, verticalalignment='top')
            else:
                plt.text(6.5, 2.5,
                         "           ".join(grouped_edges_dyHyp["{} and {}".format(tissue_1_name, tissue_2_name)]),
                         fontsize=24, verticalalignment='top')
        plt.axis('off')

        plt.text(-2, -1.7, tissue_1_name, fontsize=24, verticalalignment='top')
        plt.text(-2, -2.3, sigMediators_tissue1, fontsize=24, verticalalignment='top')
        plt.text(-2, -3.7, tissue_2_name, fontsize=24, verticalalignment='top')
        plt.text(-2, -4.3, sigMediators_tissue2, fontsize=24, verticalalignment='top')

        return figure

    def runDyNAandDyHyp(self, threshold_dyNA, std_dev_dyHyp, condition, tissue_column, title_prefix):
        #Checks for necessary user variables. Returns 0 to let user input values and try again.
        if self.errorChecks(condition, tissue_column):
            return 0
        cur_condition_data = self.readFile(condition)
        firstCytLoc = self.findFirstCyt(cur_condition_data,tissue_column)
        #"Loop to run all functions"
        all_times_str,all_times_num = self.sortTime(cur_condition_data)
        dyNA_network_complexity_all = {}
        table_rate_of_change = pandas.DataFrame()
        # Trying to get groups of connected organs
        #cur_cond_HG = hypernetx.Hypergraph.from_incidence_dataframe(cur_condition_data)
        # List of all tissues, does not include plasma
        list_organs = list(set(cur_condition_data[tissue_column]))
        list_organs.remove('Plasma')
        # Loop through time interval 0-7d
        for n in range(len(all_times_num)-1):
            cur_times_str = all_times_str[n:n + 2]
            cur_times_num = all_times_num[n:n + 2]
            # Get significant mediators with time and each other for plasma
            rel_data_plasma = cur_condition_data.loc[cur_condition_data[tissue_column] == "Plasma"]
            correl_with_time_plasma = self.get_rate_of_change_inflammatory_mediator_and_TIME(
                rel_data_plasma, cur_times_num,firstCytLoc).rename(index={0:'Plasma'})
            correl_with_other_mediators_plasma = self.get_correlation_matrix_withMEDIATORS_individual_tissue(
                rel_data_plasma,cur_times_num,firstCytLoc).rename(index={0:'Plasma'})
            pos_mediators_time_plasma, neg_mediators_time_plasma = self.get_significant_mediators_withTIME(
                correl_with_time_plasma,std_dev_dyHyp)
            pos_mediators_other_plasma, neg_mediators_other_plasma = self.get_significant_mediators_withEACHOTHER(
                correl_with_other_mediators_plasma, threshold_dyNA)
            dyNA_network_complexity_plasma = self.get_dyNA_network_complexity(pos_mediators_other_plasma,
                                                                         neg_mediators_other_plasma)
            dict_dyNA_network_complexity = {}
            dict_dyNA_network_complexity["Plasma"] = dyNA_network_complexity_plasma
            table_rate_of_change = pandas.concat([table_rate_of_change, correl_with_time_plasma])
            for j in range(0, len(list_organs)):
                rel_data_cur_organ = cur_condition_data.loc[cur_condition_data[tissue_column] == list_organs[j]]
                correl_with_time_cur_organ = self.get_rate_of_change_inflammatory_mediator_and_TIME(
                    rel_data_cur_organ,cur_times_num,firstCytLoc).rename(index={0:list_organs[j]})
                correl_with_other_mediators_cur_organ = self.get_correlation_matrix_withMEDIATORS_individual_tissue(
                    rel_data_cur_organ, cur_times_num,firstCytLoc).rename(index={0:list_organs[j]})
                pos_mediators_with_time_cur_organ, neg_mediators_with_time_cur_organ = self.get_significant_mediators_withTIME(
                    correl_with_time_cur_organ, std_dev_dyHyp)
                pos_mediators_with_other_cur_organ,neg_mediators_with_other_cur_organ = self.get_significant_mediators_withEACHOTHER(
                    correl_with_other_mediators_cur_organ, threshold_dyNA)
                group_edges_dyHyp_pos = self.group_edges_dyHyp(pos_mediators_time_plasma,
                                                          pos_mediators_with_time_cur_organ, "Plasma", list_organs[j])
                group_edges_dyHyp_neg = self.group_edges_dyHyp(neg_mediators_time_plasma,
                                                          neg_mediators_with_time_cur_organ, "Plasma", list_organs[j])
                dyNA_network_complexity_cur_organ = self.get_dyNA_network_complexity(pos_mediators_with_other_cur_organ,
                                                                                neg_mediators_with_other_cur_organ)
                dict_dyNA_network_complexity[list_organs[j]] = dyNA_network_complexity_cur_organ
                table_rate_of_change = pandas.concat([table_rate_of_change, correl_with_time_cur_organ])
                fig_pos = plt.figure(figsize=(18, 18))
                gs = fig_pos.add_gridspec(nrows=2, ncols=1, hspace=0.5, wspace=1.5)
                self.visual_graph(fig_pos, gs, group_edges_dyHyp_pos, pos_mediators_other_plasma,
                             pos_mediators_with_other_cur_organ, "Plasma", list_organs[j])
                title_pos = "{} - Positive Rate of Change {} - {}".format(condition, cur_times_str[0], cur_times_str[1])
                #fig_pos.suptitle(title_pos, size=40)
                fig_pos.savefig(title_prefix+ " " + title_pos+ " " + list_organs[j] + ".png",
                                bbox_inches="tight")

                fig_neg = plt.figure(figsize=(18, 18))
                gs = fig_neg.add_gridspec(nrows=2, ncols=1, hspace=0.5, wspace=1.5)
                self.visual_graph(fig_neg, gs, group_edges_dyHyp_neg, neg_mediators_other_plasma,
                             neg_mediators_with_other_cur_organ, "Plasma", list_organs[j])
                title_neg = "{} - Negative Rate of Change {} - {}".format(condition, cur_times_str[0], cur_times_str[1])
                #fig_neg.suptitle(title_neg, size=40)
                fig_neg.savefig(title_prefix+ " " + title_neg + " " + list_organs[j] + ".png",
                                bbox_inches="tight")
                if j < 1:
                    pos_network_plasma = self.draw_intratissue_network(
                        '{} Plasma Pos Network {}-{}'.format(title_prefix,cur_times_str[0],cur_times_str[1]),
                                                                       pos_mediators_other_plasma)
                    neg_network_plasma = self.draw_intratissue_network(
                        '{} Plasma Neg Network {}-{}'.format(title_prefix,cur_times_str[0],cur_times_str[1]),
                                                                       neg_mediators_other_plasma)
                pos_network_organ = self.draw_intratissue_network(
                    '{} {} Pos Network {}-{}'.format(title_prefix, list_organs[j], cur_times_str[0], cur_times_str[1]),
                    pos_mediators_with_other_cur_organ)
                neg_network_organ = self.draw_intratissue_network(
                    '{} {} Neg Network {}-{}'.format(title_prefix, list_organs[j], cur_times_str[0], cur_times_str[1]),
                    neg_mediators_with_other_cur_organ)
            dyNA_network_complexity_all["{} - {}".format(cur_times_str[0], cur_times_str[1])] = dict_dyNA_network_complexity
        dyNA_network_complexity_table = pandas.DataFrame(dyNA_network_complexity_all)
        writer = pandas.ExcelWriter('{} {} DyNA + Rate of Change.xlsx'.format(title_prefix, condition), engine='xlsxwriter')
        dyNA_network_complexity_table.to_excel(writer,"{}_DyNA_Complexity".format(condition))
        table_rate_of_change = table_rate_of_change.set_axis((['Plasma']+list_organs), axis='index')
        table_rate_of_change.to_excel("{}_Rate_of_Change.xlsx".format(condition))
        writer.close()

        self.close()

    def DyHyp_Network_Complexity(self, std_dev_dyHyp, condition, tissue_column, title_prefix):
        # Checks for necessary user variables. Returns 0 to let user input values and try again.
        if self.errorChecks(condition, tissue_column):
            return 0
        cur_condition_data = self.readFile(condition)
        firstCytLoc = self.findFirstCyt(cur_condition_data, tissue_column)
        #"Loop to run all functions"
        all_times_str,all_times_num = self.sortTime(cur_condition_data)
        list_organs = list(set(cur_condition_data[tissue_column]))
        # Loop through time interval 0-7d
        writer = pandas.ExcelWriter('{} {} DyHyp Network Complexity.xlsx'.format(title_prefix, condition),
                                    engine='xlsxwriter')
        for n in range(len(all_times_strt)-1):
            cur_times_str = all_times_str[n:n+2]
            cur_times_num = all_times_num[n:n+2]
            dict_pos_mediators = {}
            dict_neg_mediators = {}
            interval = cur_times_str[n] + '-' + cur_times_str[n+1]
            for j in range(len(list_organs)):
                rel_data_cur_organ = cur_condition_data.loc[cur_condition_data[tissue_column] == list_organs[j]]
                correl_with_time_cur_organ = self.get_rate_of_change_inflammatory_mediator_and_TIME(rel_data_cur_organ,
                                                                                               cur_times_num,firstCytLoc)
                dict_pos_mediators[list_organs[j]],dict_neg_mediators[list_organs[j]] = self.get_significant_mediators_withTIME(
                    correl_with_time_cur_organ,std_dev_dyHyp)
            table_pos_mediators = pandas.DataFrame(dict([(k, pandas.Series(v)) for k, v in dict_pos_mediators.items()]),dtype=object)
            table_neg_mediators = pandas.DataFrame(dict([(k, pandas.Series(v)) for k, v in dict_neg_mediators.items()]),dtype=object)
            table_pos_mediators.to_excel(writer, "{}_{}_PosCyts.xlsx".format(condition,interval), engine='xlsxwriter',
                                         index=False)
            table_neg_mediators.to_excel(writer, "{}_{}_NegCyts.xlsx".format(condition,interval), engine='xlsxwriter',
                                         index=False)
        writer.close()

        self.close()

form = Form()
form.show()
app.exec_()