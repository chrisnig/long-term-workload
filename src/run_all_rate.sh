#!/bin/bash

for i in "0" "0.1" "0.2" "0.3" "0.4" "0.5" "0.6" "0.7" "0.8" "0.9" "1"; do
	echo "Generating for target preference rate ${i}"
	python3 generate_requests.py ../data/data_generated_0.8 ../data/data_generated_rate_${i} --on_probability ${i} --off_probability 0
	cp ../data/data_generated_0.8/*.cmpl ../data/data_generated_rate_${i}
	echo "Solving for target preference rate ${i}"
	python3 solve_all.py ../data/data_generated_rate_${i} ../results/output_generated_rate_${i}
	python3 eval_solver_runs.py ../data/data_generated_rate_${i} ../results/output_generated_rate_${i} ../results/results_generated_rate_${i}.xlsx
done
