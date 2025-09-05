from api import *
import seaborn as sns
import atexit

from helpers.pd import check_cache_dir

# 🎨 nice aesthetics
sns.set_theme(style="ticks", palette="viridis")
check_cache_dir()
atexit.register(check_cache_dir)
