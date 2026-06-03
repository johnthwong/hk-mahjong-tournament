# hong kong mahjong changes:
- No Uma. remove related menus and buttons.
- Points are generally lower, like in the hundreds at most. So no need to display in thousands.
- Rephrase tiebreaker rules.
- Faan table setting. minimum and maximum faan. A field for how much the minimum faan is worth in points. the other faan should be based on the minimum faan's points, scaled either half-spicy or full-spicy (see https://en.wikipedia.org/wiki/Hong_Kong_mahjong_scoring_rules#Faan-to-score_table). admin should be to create a custom table too. the table should show the Faan (one row for each integer between by min and max faan), Points, and Self-pick points.
- A self-pick multipler field (default is 1.5x)
- penalty field for how many points to penalize for "False Win"
- third UI on the user end to see a Faan table. it should look something like this (but without the formula column)
Faan	Points	Formula
0	1	
2
0
=
1
{\displaystyle 2^{0}=1}
1	2	
2
1
=
2
{\displaystyle 2^{1}=2}
2	4	
2
2
=
4
{\displaystyle 2^{2}=4}
3	8	
2
3
=
8
{\displaystyle 2^{3}=8}
4	16	
2
4
=
16
{\displaystyle 2^{4}=16}
5	24	
1.5
×
2
4
=
24
{\displaystyle 1.5\times 2^{4}=24}
6	32	
2
5
=
32
{\displaystyle 2^{5}=32}
7	48	
1.5
×
2
5
=
48
{\displaystyle 1.5\times 2^{5}=48}
8	64	
2
6
=
64
{\displaystyle 2^{6}=64}
9	96	
1.5
×
2
6
=
96
{\displaystyle 1.5\times 2^{6}=96}
10	128	
2
7
=
128
{\displaystyle 2^{7}=128}
11	192	
1.5
×
2
7
=
192
{\displaystyle 1.5\times 2^{7}=192}
12	256	
2
8
=
256
{\displaystyle 2^{8}=256}
13	384
- no leftover field in the scoring UI.

# security:
- add a passcode gate to the admin portal (the /exec?portal=admin URL is otherwise usable by anyone with the link, since the web app runs as the owner with anonymous access).
