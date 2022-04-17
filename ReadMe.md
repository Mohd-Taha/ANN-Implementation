Implementation of Artificial Neural Network

Approach Used:      
    Independent learning using threshold error = 0.0001 on each input and then using average of all again checking (if absolute error not less than threshold) then again learning independent....
    process repeated again and again, untill condition satisfied.

Reason:             
    Using this approach I can paas inputs and target to different Threads which will boost the process.

Learning Rate:      
    Starting Rate =     1,
    After Each EPOCH =  previous * 0.95,
    Untill previous greater than 0.4 (constant after this).

File:               
    Intermediate results file are generated after each EPOCH, one can check that the weights and biases are converging.

Stopping Criteria:  
    Average of current weights and biases and previous ones are same.

Values in annDataset.xlsx are not normalized / divided all by 15000.
