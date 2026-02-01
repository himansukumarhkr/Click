import os

class ToonConfig:
    @staticmethod
    def load(filepath):
        config = {}
        if not os.path.exists(filepath): return config
        try:
            with open(filepath, 'r') as f:
                for line in f:
                    if ':' in line:
                        k, v = line.split(':', 1)
                        config[k.strip()] = v.strip() == 'True' if v.strip() in ['True', 'False'] else v.strip()
        except:
            pass
        return config

    @staticmethod
    def save(filepath, data):
        try:
            with open(filepath, 'w') as f:
                for k, v in data.items(): f.write(f"{k}: {v}\n")
        except:
            pass