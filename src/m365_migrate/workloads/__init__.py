"""Migration workloads (users, mailboxes, files, ...).

Only the Users/Identities workload is implemented today; it is the prerequisite
for the others, which key off the source→target user mapping it produces.
"""
