import subprocess
import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
script_path = os.path.join(scratch_dir, "process_chain_runner.py")

chains = ["win", "rural", "urban"]

def main():
    t_start = time.time()
    print("==================================================", flush=True)
    print("LAUNCHING ALL CHAINS IN PARALLEL...", flush=True)
    print("==================================================", flush=True)
    
    processes = {}
    log_files = {}
    
    for chain in chains:
        log_path = os.path.join(scratch_dir, f"runner_{chain}.log")
        log_file = open(log_path, "w", encoding="utf-8")
        log_files[chain] = log_file
        
        print(f"Starting process for chain: {chain.upper()} (logging to {log_path})...", flush=True)
        # Launch process in background
        proc = subprocess.Popen(
            [sys.executable, "-u", script_path, chain],
            stdout=log_file,
            stderr=subprocess.STDOUT
        )
        processes[chain] = proc
        
    print("\nWaiting for all processes to finish...", flush=True)
    
    # Active monitoring loop
    active_chains = list(chains)
    completed_chains = []
    
    while active_chains:
        for chain in list(active_chains):
            proc = processes[chain]
            ret = proc.poll()
            if ret is not None:
                active_chains.remove(chain)
                completed_chains.append(chain)
                log_files[chain].close()
                if ret == 0:
                    print(f"  [SUCCESS] Chain {chain.upper()} finished successfully in {time.time()-t_start:.1f}s!", flush=True)
                else:
                    print(f"  [ERROR] Chain {chain.upper()} failed with exit code {ret}! Check runner_{chain}.log for details.", flush=True)
        time.sleep(5)
        
    print("\n==================================================", flush=True)
    print(f"ALL PARALLEL PIPELINES COMPLETE in {time.time()-t_start:.1f}s!", flush=True)
    print("==================================================", flush=True)

if __name__ == '__main__':
    main()
