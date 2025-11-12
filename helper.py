from colorama import Fore
import traceback


def dbg(*values, sep=' ', end='\n', file=None, start=True):
    line = traceback.extract_stack()[-2].lineno

    if start:
        print(f'({Fore.RED}{line}{Fore.RESET})', *values, sep=sep, end=end, file=file)
    else:
        print(*values, f'({Fore.RED}{line}{Fore.RESET})', sep=sep, end=end, file=file)
