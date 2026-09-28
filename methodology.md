# Methodology, Project 2

## Governed taxonomy
The operator's description is authoritative. Eight governed reasons map to a code, a planned or unplanned category, a Six Big Losses class and an OEE pillar. The stored codes agree with descriptions on only about one record in eight, which is the level expected by chance.

## Time weighted OEE
For each record: operating time = planned time x A; net run time = operating time x P; fully productive time = net run time x Q. For any group, A, P and Q are ratios of summed times and OEE = summed fully productive time over summed planned time, which keeps OEE = A x P x Q exact at every level of aggregation.

## Pareto and fingerprints
Downtime minutes are ranked by reason, by loss class and by line and reason combination. Fingerprints are each line's share of downtime by reason; distance is half the sum of absolute share differences from the network. A chi square test of independence on event counts tests whether lines differ.

## Shift variance detector
Weekly time weighted OEE per line and shift is compared with the mean and standard deviation of the previous 13 weeks. A z score beyond the control limit (default 2) raises a flag; weeks with fewer than the minimum records are not scored.

## Validation
36 line months reproduce a pandas calculation exactly; the OEE identity holds; every record carries a governed code; a 61,000 event copy was timed for scale.
