import yaml

class XPathLoader:
    def __init__(self, yaml_file, main_thread_name):
        self.xpaths = self._load_xpaths(yaml_file)
        self.main_thread_name = main_thread_name

    def _load_xpaths(self, yaml_file):
        with open(yaml_file, 'r') as file:
            config = yaml.safe_load(file)
        return config.get(self.main_thread_name, {})

# Usage
loader = XPathLoader('orizon.yaml', main_thread_name='xpaths')