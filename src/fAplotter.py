# %% Import the required Python packages:

import pandas as pd
from scipy.interpolate import interp1d
from scipy.optimize import root_scalar
from scipy.integrate import quad

# %% Initiate Synergy (API):

import win32com.client
Synergy = win32com.client.Dispatch('synergy.Synergy')
Synergy.SetUnits('Metric')

# %% Define the functions to be used later:


def get_nodes_from_AMI_selection_list():
    # Get the current entries of the AMI selection list:
    # Synergy (in AMI): Mesh -> Selection -> Selection list.
    SD = Synergy.StudyDoc
    SelectList = SD.Selection
    list_size = SelectList.Size
    l_nodes = []
    for i in range(list_size):
        Ent = SelectList.Entity(i)
        list_entry = Ent.ConvertToString
        if list_entry[0] == 'N':
            node = list_entry[1:]  # 'N123' --> '123'
            l_nodes.append(int(node))
    return l_nodes


def get_AMI_nodal_results(idname='Temperature'):
    # Create a DoubleArray object for the time steps (s):
    IndpValues = Synergy.CreateDoubleArray
    PlotManager = Synergy.PlotManager
    ResultID = PlotManager.FindDatasetIDByName(idname)
    PlotManager.GetIndpValues(ResultID, IndpValues)
    l_timesteps = sorted(list(IndpValues.ToVBSArray))  # Pyhton list - time (s)

    dict_results = {}  # Python dictionary to temporarily store nodal results
    # Retrieve nodal results for each time step:
    for t in l_timesteps:
        # Create a DoubleArray object for the (current) time (s):
        Ind = Synergy.CreateDoubleArray
        Ind.AddDouble(t)

        # Create an IntegerArray object for the node numbers,
        # and a DoubleArray object for the result values:
        I_NodeNb = Synergy.CreateIntegerArray
        D_NodalResults = Synergy.CreateDoubleArray
        PlotManager.GetScalarData(ResultID, Ind, I_NodeNb, D_NodalResults, )

        l_nodes = list(I_NodeNb.ToVBSArray)  # Python list - nodes
        # nodal results (Pyhton list):
        l_nodalresults = list(D_NodalResults.ToVBSArray)

        # Very high float values are assigned to sections (nodes) that have not
        # been filled (at the given time). Those values are replaced with None.
        if max(l_nodalresults) > 1e30:
            l_nodalresults = [i if i < 1e30 else None for i in l_nodalresults]

        # Temporarily store the current time step results in a Pandas Series:
        ser = pd.Series(data=l_nodalresults, index=l_nodes, name=t)
        dict_results[t] = ser

    # "Reset" PlotManager, to avoid issues when using it again:
    PlotManager = None

    # Store the nodal results in a Pandas DataFrame, with
    # time steps (s) as index , and nodes as columns:
    df_results = pd.DataFrame(data=dict_results).T

    return df_results


def get_factors(ser_tT, T_crit):
    # Remove the rows which contain only None
    # ("nodes that are not yet filled"):
    ser_tT = ser_tT.dropna(axis=0, how='all',)

    t_melt_contact = ser_tT.index.min()  # time at initial melt contact (s)
    t_cycle_end = ser_tT.index.max()  # time at the end of the cycle (s)

    # Create the interpolation function T(t)-T_crit:
    f_T = interp1d(ser_tT.index, ser_tT-T_crit,)  # (°C)

    # Solve for T(t)-T_crit=0:
    sol = root_scalar(f_T, bracket=[t_melt_contact, t_cycle_end],)
    # Time within the cycle, when the melt has cooled down to T_crit:
    t_cooled = sol.root  # (s)

    ft = t_cooled - t_melt_contact  # Time (s) the melt is above T_crit

    # Integrate int_{t_melt_contact}^t_cooled T(t)-T_crit dt:
    def f_integrate(t, fun): return fun(t)
    fA = quad(f_integrate, t_melt_contact, t_cooled,
              args=(f_T), full_output=True)[0]  # (K*s)

    return ft, fA  # (s), (K*s)


def create_AMI_single_scalar_plot(ser_result,
                                  plot_title='Custom AMI plot', unit_plt='s'):
    # Initiate custom AMI contour plot:
    PlotManager = Synergy.PlotManager
    CreateUserPlot = PlotManager.CreateUserPlot
    CreateUserPlot.SetDataType('NDDT')
    CreateUserPlot.SetName(plot_title)  # Result name

    # This contour plot has no independent variable (such as time in s):
    CreateUserPlot.SetIndpName('No independent available')
    CreateUserPlot.SetIndpUnitName('')
    # Set name and unit of dependent variable:
    CreateUserPlot.SetDeptName('fA')
    CreateUserPlot.SetDeptUnitName('K*s')

    # Create an IntegerArray object for the node numbers,
    # and a DoubleArray object for the result values:
    I_NodeNb = Synergy.CreateIntegerArray
    D_NodalResults = Synergy.CreateDoubleArray
    # Assign an arbitrary Python float value to the independent variable:
    aIndpValue = 0.

    # Add the node number and result values to the corresponding arrays:
    for node, value in ser_result.items():
        if value is None:  # Skip rows with None:
            continue
        I_NodeNb.AddInteger(int(node))
        D_NodalResults.AddDouble(float(value))

    # Creates the AMI contour plot:
    CreateUserPlot.AddScalarData(aIndpValue, I_NodeNb, D_NodalResults)
    CreateUserPlot.Build
    PlotManager = None

    # Save the (active) AMI Study file:
    SD = Synergy.StudyDoc
    SD.Save

    return


# %% Run script:
if __name__ == '__main__':
    # Glass transition temperature of PC in °C
    Tg = float(input("Enter Tg in °C: ") or 140)
    print(f'Tg={Tg} °C')
    
    # Nodes for which ft- and fA-factor will be calculated (Python list):
    # V1: Get nodes from the AMI Selection list
    #     (Select nodes by mouse -> Mesh -> Selection -> Selection list:
    l_nodes = get_nodes_from_AMI_selection_list()
    # V2: Alternatively, nodes can be provided manually:
    # str_nodes = 'N491038 N491099 N490080 N489301 N490081 N491039 N491037'
    # l_nodes = [int(i) for i in str_nodes.replace('N', '').split(' ')]

    # Obtain (nodal) temperature results from AMI (Pandas DataFrame):
    df_nodalT = get_AMI_nodal_results(idname='Temperature')
    # df_nodalT.to_pickle('nodal_T_results.pkl')
    # df_nodalT = pd.read_pickle('nodal_T_results.pkl')

    # Extract results for nodes in the list (Pandas DataFrame):
    df_nodalT = df_nodalT.loc[:, l_nodes]

    # Calculate the ft- and fA-factors (Pandas DataFrame):
    def func_ft(ser): return get_factors(ser, Tg)
    df_factors = df_nodalT.apply(func_ft, axis=0, result_type='expand')

    # ft-factors and fA-factors for all nodes in the list (Pandas Series):
    ser_ft = df_factors.iloc[0, :]  # (s)
    ser_fA = df_factors.iloc[1, :]  # (K*s)

    # Create an AMI custom contour plot showing the local fA-factors:
    plot_title = f'fA-factor plot (average: {ser_fA.mean():.2f} K*s)'
    create_AMI_single_scalar_plot(ser_fA, plot_title=plot_title,
                                  unit_plt='K*s')
    input(f"Script completed. The averaged fA={ser_fA.mean():.2f} K*s\nPress any key to close.")
