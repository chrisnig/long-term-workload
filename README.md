Long-term fairness for preference fulfillment of hospital physicians
====================================================================

This is an implementation to assist with experiments presented in the paper "Hospital physicians can't get no satisfaction - An indicator for fairness in preference fulfillment on duty schedules".

Directory structure
-------------------

This repository is structured as follows.

* ``data`` contains the data used for the experiments in the paper. For privacy reasons, we do not provide real-life data but only generated data.
    * ``data/data_generated_0.8`` contains data generated based on the real-life data without restricting the conflict rate and a preference rate of 80%. We include this data because it is the base data we used to generate our data with a target conflict rate.
    * ``data/data_generated_conf_{x}`` contain data with a conflict rate of *x*. This is the data used for our experiments in the paper.
* ``results`` contains the results obtained by experiments in the paper. ``results/output_generated`` and ``results/results_generated.xlsx`` contain results for the generated data. For privacy reasons, we do not provide detailed results for real-life data, instead we just provide aggregated real-life results in ``results/results_real.xlsx``.
* ``src`` contains all the scripts required to run the experiments described in the paper. See below for a step-by-step description how to run them.

Requirements
------------

* [Python 3.5](https://www.python.org/) (might work with earlier version of Python 3, but we did not test it)
* openpyxl (install via ``pip``)
* [CMPL 1.11.0](https://coliop.org)
* [IBM ILOG CPLEX](https://www.ibm.com/us-en/marketplace/ibm-ilog-cplex) (or another solver compatible with CMPL, but we only tested with CPLEX)

The ``cmpl`` executable must be in the PATH variable for the software to find it.

We recommend using Linux as we provide a bash script and CMPL support for Linux is superior. Running on Windows is untested.

Running
-------

To reproduce our experiments, we provide a bash script that performs all required steps. It first generates data with a target conflict rate for each conflict rate between 0% and 100% with 10% increments. Afterwards, it runs all updating strategies for each conflict rate and evaluates the results into Excel sheets. This is the easiest way to reproduce our experiments.

Before invoking the script, you should clear all of our generated data and results as follows.
* Delete everything in the ``results`` directory, but not the directory itself.
* Delete all directories named ``data/data_generated_conf_*``, but make sure not to delete ``data/data_generated_0.8`` as this directory will be required for data generation.

After you have performed the deletion, you can change to the ``src`` directory and invoke ``run_all_conf.sh``. After the script has finished, you can find all the results in the ``results`` directory.

Alternatively, you can also go through all steps manually. The process is described below. We show the process for a target conflict rate of 80%. If you want to compare different conflict probabilities, the process needs to be repeated.

1. **Generate data for physicians.** This step is optional as we supply generated data in the ``data`` directory which is the data we ran our experiments on. Skip this step if you want to use our generated data.

    Our software assumes the data in the form of CDAT files. These files are read by the CMPL software. We expect each CDAT file to have the same file name. The files should each be stored in a different subdirectory. In our case, we name the subdirectories with the starting date of the respective roster and the number of weeks the roster should span. The following parameters need to be supplied to the script to generate data: 
    * input directory which contains subdirectories with existing data that should be modified. You can use our supplied data for this purpose.
    * output directory where subdirectories for the generated data will be created
    * number of physicians
    * number of duties
    * percentage of multi-skilled physicians (i.e., physicians with two qualifications instead of one)
    * (optional) --on_probability probability for each day that a physician has a preference for a duty
    * (optional) --off_probability probability for each day that a physician has a preference for not being assigned to a duty
    * (optional) --filename file name of the CDAT file. Note that our solver script expects this to be "transformed.cdat".
    * (optional) --conflict_probability likelihood in percent that a duty request is in conflict with at least one other request

    A sample invocation of the command would look as follows: ``python src/generate_params.py data/data_generated_0.8 data/my_generated_data 85 6 0.2 --on_probability 0.8 --off_probability 0 --filename transformed.cdat --conflict_probability 0.8``

    Alternatively, you may choose just to generate new requests and keep the data for physicians, duties, and qualifications. For this purpose, we provide the script ``src/generate_requests.py`` which takes the following parameters:

    * input directory which contains subdirectories with existing data that should be modified. You can use our supplied data for this purpose.
    * output directory where subdirectories for the generated data will be created
    * (optional) --on_probability probability for each day that a physician has a preference for a duty
    * (optional) --off_probability probability for each day that a physician has a preference for not being assigned to a duty
    * (optional) --filename file name of the CDAT file. Note that our solver script expects this to be "transformed.cdat".
    * (optional) --conflict_probability likelihood in percent that a duty request is in conflict with at least one other request
    
    If you want to compare results for different conflict rates, we recommend generating the data only once and then using the request generation tool to generate different instances of requests. The advantage of this approach is that all of your data instances will have the same physicians with the same qualifications and only the requests will be different.

2. **Copy the model to your data directory.** If you are using our supplied data, skip this step as the models are already in the data directory.

    Our solver run script assumes the models for the different strategies to be named ``ltf-unfair.cmpl`` (C strategy), ``ltf-fair-linear.cmpl`` (ESA strategy), and ``ltf-fair-nonlinear.cmpl`` (ESD strategy). You can copy these files from our supplied data directory and put them directly in your data directory (not in one of the subdirectories). For an example, see the supplied ``data/data_generated_conf_0.8`` directory.

3. **Run the solver on the data.**

    We provide a script which automatically invokes the solver for each instance and also takes care of updating the historic satisfaction  as described in the paper. Note that the execution of this command may take a while. It doesn't output any messages, so don't worry if it looks like nothing is happening. Check your computer's CPU usage and task list to see whether it is still running.

    Our solver run script takes two parameters:
    * input directory which contains the model files and subdirectories with the data. Note that there should be a file called ``transformed.cdat`` in each subdirectory containing the data for the instance.
    * output directory where the solution will be written

    A sample invocation of the command looks as follows: ``python src/solve_all.py data/my_generated_data results/my_new_results``

4. **Evaluate the results in Excel format.**

    The solver has solved all the models, so now we can aggregate the data. The ``eval_solver_runs.py`` script creates an Excel file that provides an overview of the data and also creates the nifty graphs shown in the paper.

    The evaluation script takes three parameters:
    * the solver input directory containing the data
    * the solver output directory containing the solutions
    * the file name of the results file that should be written

    A sample invocation of the command looks as follows: ``python src/eval_solver_runs.py data/my_generated_data results/my_new_results results/my_new_results/my_evaluation_results.xlsx``
 
    If everything worked out, you should now have an Excel spreadsheet at the location you specified. It will contain the charts on the first tab and the values for the performance indicators on the second tab as the bottom of the leftmost table.
