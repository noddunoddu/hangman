{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 6,
   "metadata": {},
   "outputs": [],
   "source": [
    "import xlwings as xw\n",
    "import pandas as pd"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "metadata": {},
   "outputs": [],
   "source": [
    "wb = xw.Book()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 4,
   "metadata": {},
   "outputs": [],
   "source": [
    "wb2 = xw.Book(r\"C:\\Users\\respe\\OneDrive\\デスクトップ\\読み取りデータ.csv\")"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 45,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "[16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33]\n"
     ]
    },
    {
     "data": {
      "text/plain": [
       "16"
      ]
     },
     "execution_count": 45,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "_df = pd.read_csv(r\"C:\\Users\\respe\\OneDrive\\デスクトップ\\読み取りデータ.csv\",encoding = 'cp932')\n",
    "_df_row = _df.index[_df['通常・延滞']  == '延滞'].tolist()\n",
    "print(_df_row)\n",
    "haritsukeretu = _df_row[0]\n",
    "haritsukeretu"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 38,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<bound method Sheet.activate of <Sheet [読み取りデータ.csv]読み取りデータ>>"
      ]
     },
     "execution_count": 38,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "ws2 = wb2.sheets(1)\n",
    "ws2.activate"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 46,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "[['通常', 1.0, 100.0],\n",
       " ['通常', 2.0, 500.0],\n",
       " ['通常', 3.0, 10500.0],\n",
       " ['通常', 4.0, 20500.0],\n",
       " ['通常', 5.0, 30500.0],\n",
       " ['通常', 6.0, 40500.0],\n",
       " ['通常', 7.0, 50500.0],\n",
       " ['通常', 8.0, 60500.0],\n",
       " ['通常', 9.0, 70500.0],\n",
       " ['通常', 10.0, 80500.0],\n",
       " ['通常', 11.0, 90500.0],\n",
       " ['通常', 12.0, 100500.0],\n",
       " ['通常', 13.0, 110500.0],\n",
       " ['通常', 14.0, 120500.0],\n",
       " ['通常', 15.0, 130500.0]]"
      ]
     },
     "execution_count": 46,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "haritsuke = ws2.range((2,1),(str(haritsukeretu),3) ).value\n",
    "haritsuke"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 47,
   "metadata": {},
   "outputs": [],
   "source": [
    "wb.sheets(1).range(\"A1\").value = haritsuke"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "interpreter": {
   "hash": "1d36b02e311e5684886219c57d0abfecc2b73caa5a7e3cf50223a94787af0ca4"
  },
  "kernelspec": {
   "display_name": "Python 3.10.4 64-bit",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.10.4"
  },
  "orig_nbformat": 4
 },
 "nbformat": 4,
 "nbformat_minor": 2
}
