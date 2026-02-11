from statsmodels.stats.weightstats import DescrStatsW as ws
import polars.selectors as cs
import polars as pl
from polars import col as c
from pathlib import Path
import pandas as pd 
import pyreadstat
from functools import wraps

def set_var_labels(labels: dict) -> None:
    global var_labels
    var_labels = labels

def apply_labels(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if (var_labels is None) or (kwargs.get('labels', True) is False):
            return func(*args, **kwargs)
        else:
            _df:pl.DataFrame = func(*args, **kwargs)
            cols = [c for c in _df.columns if c in var_labels.keys()]
            return _df.with_columns(*[getattr(c(col), 'replace_strict')(var_labels[col]) for col in cols])
    return wrapper

def check_label(dictionary:dict, word:str, parent:str='') -> None:
    '''
    Searches for word either inside the dictionary keys or in its labels.
    Search is done in lowecase 
    
    Args:
        dictionary (dict) :
        word (str) :
    
    Returns:
        None
    '''
    word=word.casefold()
    for key, value in dictionary.items():
        path = f"{parent}.{key}" if parent else key

        if (word in str(key).casefold()):
            print(path, ": ", value, end='\n'*2)
            
        elif not isinstance(value, dict):
            if (word in str(value).casefold()):
                print(path, ": ", value, end='\n'*2)
        else:
            check_label(value, word, path)


def check_code_label(df:pl.DataFrame, column:str) -> pl.DataFrame:
    '''
    Check the column's labels and output a pl.DataFrame table with the numeric codes and the labels for each code
    
    Args:
        df (pl.DataFrame): DataFrame containing the variables
        column (str): Column of interest
        var_labels (dict): by default selects the global variable called var_labels

    Returns:
        pl.DataFrame with 2 columns one containing the numeric codes, and the other containing the labels
    '''
    return df.select(
                c(column).unique()
            ).sort(
                by=column
            ).with_columns(
                c(column)
                .replace_strict(var_labels[column], default=c(column))
                .alias('labels')
            )

def w_stat(df, col, factor):
    return ws(df[col], df[factor])

def read_file(file:Path, output_format:str='polars', metadata_only:bool=False, encoding='utf-8') -> [pl.DataFrame, dict, dict]:
    """
    Loads .sav or .dta and returns a dataframe with the metadata

    Args:
        file (Path): path_to_file
        output_format (str): Polars | Pandas
        metadata_only (bool, optional): Defaults to False.

    Returns:
        list ([pl.DataFrame | pd.DataFrame, dict, dict]): df, var_labels, col_labels
    """    
    df, meta = getattr(pyreadstat, f'read_{file.suffix[1:]}')(file, output_format=output_format, metadataonly=metadata_only)
    df:pl.DataFrame|pd.DataFrame
    df = df.rename(str.casefold)
    var_labels = {k.casefold(): v for k, v in meta.variable_value_labels.items()}
    col_labels = {k.casefold(): v for k, v in meta.column_names_to_labels.items()}
    return df, var_labels, col_labels

@apply_labels
def tab(df:pl.DataFrame, col:str, groups:list[str]=False, factor=None, decimals:int=2, **kwargs) -> pl.DataFrame:
    """
    Create a table to show counts, and proportions

    Args:
        df (pl.DataFrame):
        col (str): main aggregation
        groups (list[str]):  
        factor (_type_, optional): Observation weights in case necessary. Defaults to None.

    Returns:
        pl.DataFrame
    """    
    if not factor:
        df = df.with_columns(pl.lit(1).alias('__factor'))
        factor='__factor'
    if groups:
        if not isinstance(groups, list):
            groups=[groups]
        groups_extended = groups + [col]
        return (
            df.group_by(*groups_extended)
            .agg(c(factor).sum().alias('Nobs'))
            .sort(*groups_extended)
            .with_columns(
                c.Nobs.truediv(c.Nobs.sum()).mul(100).over(*groups).round(decimals)
                .alias('proportion')
            )
            .with_columns(
                c.proportion.cum_sum().over(*groups).round(decimals)
                .alias('cum')
            )
        )
    else:
        return (
            df.group_by(col).agg(c(factor).sum().alias('Nobs'))
            .sort(col)
            .with_columns(c.Nobs.truediv(c.Nobs.sum()).mul(100).round(decimals)
                          .alias('proportion')
            )
            .with_columns(
                c.proportion.cum_sum().round(decimals)
                .alias('cum')
            )
            
        )

@apply_labels
def compute_stat(df:pl.DataFrame, col:str, groups:str|list, stats:list[str]=['mean', 'std'], factor:str=None, decimals:int=2, **kwargs):
    """
    Compute aggregated statistics, weights are optional

    Args:
        df (pl.DataFrame): dataframe
        col (str): variable to compute statistic
        groups (str | list): aggregation groups
        stats (list[str], optional): Statistics to calculate. Defaults to ['mean', 'std'].
        factor (str, optional): Weights for each observation. Defaults to None.

    Returns:
        pl.DataFrame: Table with summary stats
    """    
    if not factor:
        df = df.with_columns(pl.lit(1).alias('__factor'))
        factor='__factor'
    if isinstance(stats, str):
        stats = [stats]
    if isinstance(groups, str):
        groups = [groups]
    if factor=='__factor':
        _df (
            df.group_by(*groups)
            .agg(*[getattr(c(col), i)().alias(i) for i in stats], c(factor).sum().alias('nobs'))
            )
    else:
        stats = stats + ['nobs']
        _df = (
            df.group_by(*groups)
            .map_groups(lambda gp: (
                wst := w_stat(gp, col, factor=factor),
                gp.select(c(groups).first(), *[pl.lit(getattr(wst, i)).alias(i) for i in stats])
                )[1]
            )
        )
        
    return _df.sort(*groups).with_columns(cs.float().round(decimals))

