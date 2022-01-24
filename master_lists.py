optout_df_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal Title", "Journal License",
                  "Author First Name", "Author Last Name", "Author Email", "Author Affiliation",
                  "Initial Institution (sent to) BPID", "Eligibility Completed", "Comments", "OASiS Status",
                  "Reason for OASiS Status Change", "Agreement", "Licence ID", "Opt-Out Reason",
                  "Monthly Reporting Status"
                  ]

palm_optout_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal License", "Author First Name",
                    "Author Last Name", "Author Email", "Eligibility Completed", "Initial Institution (sent to) BPID",
                    "OASiS Status", "Reason for OASiS Status Change", "Opt-Out Reason"
                    ]

palm_req_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal License", "Author First Name",
                 "Author Last Name", "Author Email", "Date Sent to AAS", "Date Approved",
                 "Date Rejected", "Date Forwarded", "Initial Institution (sent to) BPID", "OASiS Status",
                 "Reason for OASiS Status Change"
                 ]

aas_dashboard_cols = ["DOI", "Article Title", "Author First Name", "Author Last Name", "Journal Title",
                      "Article Type", "Approver Email", "Date Sent to AAS", "Date Approved", "Date Rejected",
                      "Date Forwarded", "Journal License", "Author Share", "Institutional Share",
                      "Reason for Full Coverage"
                      ]

complete_oa_req_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal Title", "Journal License",
                        "Author First Name", "Author Last Name", "Author Email", "Author Affiliation",
                        "Initial Institution (sent to) BPID", "Date Sent to AAS", "Date Approved", "Date Rejected",
                        "Date Forwarded", "Comments", "Approver Email", "OASiS Status",
                        "Reason for OASiS Status Change",
                        "Agreement", "Licence ID", "Final Reminder to be Sent", "Final Reminder Sent?", "Author Share",
                        "Institutional Share", "Reason for Full Coverage", "Monthly Reporting Status"]

approved_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal Title", "Journal License",
                 "Author First Name", "Author Last Name", "Author Email", "Author Affiliation",
                 "Initial Institution (sent to) BPID", "Date Sent to AAS", "Date Approved", "Date Forwarded",
                 "Comments",
                 "Approver Email", "OASiS Status", "Reason for OASiS Status Change", "Agreement", "Licence ID",
                 "Final Reminder to be Sent", "Final Reminder Sent?", "Author Share", "Institutional Share",
                 "Reason for Full Coverage", "Monthly Reporting Status"]

rejected_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal Title", "Journal License",
                 "Author First Name", "Author Last Name", "Author Email", "Author Affiliation",
                 "Initial Institution (sent to) BPID", "Date Sent to AAS", "Date Rejected", "Date Forwarded",
                 "Comments",
                 "Approver Email", "OASiS Status", "Reason for OASiS Status Change", "Agreement", "Licence ID",
                 "Final Reminder to be Sent", "Final Reminder Sent?", "Author Share", "Institutional Share",
                 "Reason for Full Coverage", "Monthly Reporting Status"]

in_verification_cols = ["DOI", "Article Title", "Article Type", "Journal ID", "Journal Title", "Journal License",
                        "Author First Name", "Author Last Name", "Author Email", "Author Affiliation",
                        "Initial Institution (sent to) BPID", "Date Sent to AAS", "Date Forwarded", "Comments",
                        "Approver Email", "OASiS Status", "Reason for OASiS Status Change", "Agreement", "Licence ID",
                        "Final Reminder to be Sent", "Final Reminder Sent?", "Author Share", "Institutional Share",
                        "Reason for Full Coverage", "Monthly Reporting Status"]

AAS_rep_columns_foal = ["DOI", "Article Title", "Author First Name", "Author Last Name", "Journal Title",
                        "Approving Institution", "Article Type", "Approver Email", "Date Sent to AAS", "Date Approved",
                        "Date Rejected", "Date Forwarded", "Journal License", "Author Share", "Institutional Share",
                        "Reason for Full Coverage"
                        ]

cancelled_cols = ["DOI", "Article Title", "Verification Status (old DOI)", "Journal ID", "OASiS Status",
                      "Reason for OASiS Status Change", "Reason for Deletion or DNR",
                      "Comments (to be filled in only if columns G-H are 'Other' or empty)", "New DOI",
                      "Verification Status (new DOI)", "Agreement", "Date", "Monthly Reporting Status"]

aas_rep_update_foal = ["Article Title", "Author First Name", "Author Last Name", "Journal Title",
                       "Approving Institution",  "Article Type", "Approver Email", "Date Sent to AAS", "Date Approved",
                       "Date Rejected", "Date Forwarded", "Journal License", "Author Share", "Institutional Share",
                       "Reason for Full Coverage"
                       ]

FOAL_columns = ["DOI", "PRS Article ID", "Unique Article ID", "Article Title", "Article Type", "Journal Title",
                "Journal License", "Author First Name", "Author Last Name", "Author Email", "Selected Affiliation",
                "Approving Institution", "Date Sent to AAS", "Date Approved", "Date Rejected", "Date Forwarded",
                "Comments", "Approver Email", "PRS Institution", "Submission Institution", "Submission Institution BPID",
                "Selected Affiliation BPID", "Email BPID", "Email Institution", "IP BPID", "IP Institution",
                "Institution Assigned at Acceptance", "Institution Assigned at Acceptance BPID", "OC Membership State",
                "Waiver Type", "Agreement Type", "Agreement", "Licence ID", "Final Reminder to be Sent",
                "Final Reminder Sent?", "Author Share", "Institutional Share", "Reason for Full Coverage",
                "Monthly Reporting Status"]

recon_report_columns = ["DOI", "PRS Article ID", "Unique Article ID", "Article Title", "Article Type", "Journal Title",
                        "Journal License", "Author First Name", "Author Last Name", "Author Email",
                        "Selected Affiliation", "Approving Institution", "Date Sent to AAS", "Date Approved",
                        "Date Rejected", "Date Forwarded", "Approver Email", "PRS Institution", "Submission Institution",
                        "Submission Institution BPID", "Selected Affiliation BPID", "Email BPID", "Email Institution",
                        "IP BPID", "IP Institution", "Institution Assigned at Acceptance",
                        "Institution Assigned at Acceptance BPID", "OC Membership State", "Waiver Type"]

foal_update_cols = ["DOI", "PRS Article ID", "Article Title", "Selected Affiliation", "PRS Institution",
                    "Journal Title", "Article Type", "Approver Email", "Date Sent to AAS", "Date Approved",
                    "Date Rejected", "Date Forwarded", "Approving Institution", "Author Email", "Journal License",
                    "Submission Institution", "Submission Institution BPID", "Selected Affiliation BPID",
                    "IP Institution", "IP BPID", "Email Institution", "Email BPID",
                    "Institution Assigned at Acceptance", "Institution Assigned at Acceptance BPID",
                    "OC Membership State", "Waiver Type", "Author First Name", "Author Last Name"]

aas_rep_cols = ["DOI", "Article Title", "Corresponding Author", "Journal Title", "Approving Institution",
                "Article Type", "Approver Email", "Date Sent to AAS", "Date Approved", "Date Rejected",
                "Date Forwarded", "License", "Inserted?", "Author Share", "Institutional Share",
                "Reason for Full Coverage"]

aas_rep_update = ["Article Title", "Corresponding Author", "Journal Title", "Approving Institution", "Article Type",
                  "Approver Email", "Date Sent to AAS", "Date Approved", "Date Rejected", "Date Forwarded", "License",
                  "Inserted?", "Author Share", "Institutional Share", "Reason for Full Coverage"
                  ]

standalone_list = ["UDR", "SU", "UMS", "CAS", "VIU", "VMU", "NIH", "WIS", "FRANCIS CRICK", "SEMMELWEIS U"]

del_cols_to_overwrite = ["Article Title", "Verification Status (old DOI)", "OASiS Status",
                         "Reason for OASiS Status Change", "Reason for Deletion or DNR",
                         "Comments (to be filled in only if columns G-H are 'Other' or empty)", "New DOI",
                         "Verification Status (new DOI)", "Agreement", "Date", "Journal ID", "Monthly Reporting Status"]

col_times_foal = ["Date Sent to AAS", "Date Approved", "Date Rejected", "Date Forwarded"]

final_del_cols = ["articledoi", "Agreement", "Journal ID", "Monthly Reporting Status"]

hybrid_errors_cols = ["DOI", "Date Sent to AAS", "Date Approved", "Date Rejected", "Comments", "Approving Institution",
                      "Article Type", "Journal License", "Author Affiliation", "OASiS Status",
                      "Reason for OASiS Status Change", "Issue", "Action to take", "sheet in file"
                      ]

foal_errors_cols = ["DOI", "PRS Article ID", "Unique Article ID", "Date Sent to AAS", "Date Approved", "Date Rejected",
                    "Comments", "Approving Institution", "Article Type", "Journal License", "Selected Affiliation",
                    "OC Membership State", "Waiver Type", "Issue", "Action to take"
                    ]

all_errors_cols = ["DOI", "PRS Article ID", "Unique Article ID", "Date Sent to AAS", "Date Approved", "Date Rejected",
                   "Comments", "Approving Institution", "Article Type", "Journal License", "Author Affiliation",
                   "Selected Affiliation", "OASiS Status", "Reason for OASiS Status Change", "OC Membership State",
                   "Waiver Type", "Issue", "Action to take", "sheet in file"
                    ]

# article/license types & final reminders

normal_article_types = ["OriginalPaper", "ReviewPaper", "BriefCommunication", "ContinuingEducation"]
deal_article_types = ["OriginalPaper", "ReviewPaper", "BriefCommunication", "BookReview", "EditorialNotes", "Report",
                      "Letter"]

all_memberships_states = ["MEMBERSHIP_FOUND", "REJECTED", "REJECTED_AND_BMC_MEMBERSHIP_FOUND", "WAIVER_INSTITUTION",
                          "WAITING_FOR_APPROVAL"]

fullyoa_bibsam = ["Swedish Research Council for Health, Working Life and Welfare (Forte)",
                  "Swedish Agency of Marine and Water Management",
                  "Swedish Radiation Safety Authority",
                  "Swedish Research Council", "Vinnova"
                  ]

reg_at = ["BIBSAM", "HAS", "QNL"]

## add Luxembourg & CNR?
fr_in_3bdays = ["UKB", "KEMÖ", "JISC", "BIBSAM", "HAS", "ICM", "DEAL", "SIKT", "MANI", "CSAL", "CRUI-CARE", "UDR", "SU",
                "IREL", "CDL", "NIH", "UMS", "VIU", "CAS"]

fr_in_5bdays = ["QNL", "FINELIB"]
approve_after2 = ["BIBSAM"]
approve_after3 = ["UKB", "JISC", "KEMÖ", "HAS", "ICM", "QNL"]
## add Luxembourg & CNR?
approve_after4 = ["DEAL", "SIKT", "MANI", "CSAL", "CARE", "UDR", "SU", "IREL", "CDL", "NIH", "UMS", "VIU", "CAS"]
approve_after5 = ["FINELIB"]

# lists for error trends

files_to_delete = ["Total no. of deleted DOIs", "Old vs. New DOI", "Reason for deletion", "Deletion by Agreement",
                   "Deletion by Journal", "Deletion by PE"]

removed = ["Found through checks", "Found through checks progress"]

horizontal_bar_chart_colours = [(0.11, 0.82, 0.25), (0.9, 0.02, 0.02), (0.8, 0.8, 0.8), (0, 0.8, 1),
                                (1, 0.67, 0), (0.84, 0.47, 1), (1, 1, 0)
                                ]

old_new_list = ["DOI", "Verification Status"]
reason_del_list = ["Reason for Deletion", "Total DOIs", "% as number"]
req_cols = ["Approved", "Rejected", "Not verified", "Opt-Out", "Ineligible for TA", "No new DOI",
            "Waiting for Author to Complete OASiS"]

palm_dates = ["Date Sent to AAS", "Date Approved", "Date Rejected", "Date Forwarded", "Eligibility Completed"]
appr_time_cols = ["Date Sent to AAS", "Date Approved", "Date Forwarded"]
rej_time_cols = ["Date Sent to AAS", "Date Rejected", "Date Forwarded"]
unver_time_cols = ["Date Sent to AAS", "Date Forwarded"]
optout_time_cols = ["Eligibility Completed"]
dnr_time_cols = ["Date"]
