# DAGMaintenance
This is a set of scripts to assist with automatic patching of Exchange 2019 DAG Members
The scripts are designed to be run on each DAG Member, and they will put the server into maintenance mode, patch the server, reboot if required and take it out of maintenance mode.

These scripts require running in an Exchange Powershell window in Administrator mode (to address the Cluster), and also PSWindowsUpdate to automate the patching.

No Parameters are needed, as the scripts pull all necessary variables from the environment. If you have a non-standard setup that the scripts can't accommodate. Tell me what it is.
I'm not promising to change the scripts to deal with your edge case, but I will be curious to know how you have two computernames for one machine, or get anything done without a well-configured DNS environment.

These scripts are currently on v4.2, there is some error-checking but not as much as I would like. Currently the scripts assume the Cluster commands and Redirect-Message command execute successfully. This will be improved.
