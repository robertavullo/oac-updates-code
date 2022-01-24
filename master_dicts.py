aas_rep_dict = {"article title": "Article Title",
                "journal title": "Journal Title",
                "article type": "Article Type",
                "approver": "Approver Email",
                "approval requested date": "Date Sent to AAS",
                "approval date": "Date Approved",
                "rejection date": "Date Rejected",
                "forward to date": "Date Forwarded",
                "license type": "Journal License",
                "AuthorShare": "Author Share",
                "InstitutionalShare": "Institutional Share",
                "Reason for Full coverage": "Reason for Full Coverage",
                "journal number": "Journal ID"
                }

AAS_rep_dict_foal = {"article title": "Article Title",
                     "corresponding author name": "Corresponding Author",
                     "journal title": "Journal Title",
                     "membership institute": "Approving Institution",
                     "article type": "Article Type",
                     "approver": "Approver Email",
                     "approval requested date": "Date Sent to AAS",
                     "approval date": "Date Approved",
                     "rejection date": "Date Rejected",
                     "forward to date": "Date Forwarded",
                     "license type": "Journal License",
                     "AuthorShare": "Author Share",
                     "InstitutionalShare": "Institutional Share",
                     "Reason for Full coverage": "Reason for Full Coverage",
                     "Eligibility completed": "Eligibility Completed"
                     }

palm_col_dict = {"Entered custom affiliation": "Custom Affiliation",
                 "Initial institution (sent to) BPID": "Initial Institution (sent to) BPID",
                 "Article title": "Article Title",
                 "Approval flow started": "Date Sent to AAS",
                 "Approved": "Date Approved",
                 "Rejected": "Date Rejected",
                 "Forwarded": "Date Forwarded",
                 "Author email": "Author Email",
                 "Licence type": "Journal License",
                 "Author First name": "Author First Name",
                 "Author Last name": "Author Last Name",
                 "Article type": "Article Type",
                 "Opt out reason": "Opt-Out Reason",
                 "OASiS status": "OASiS Status",
                 "Reason for abort/cancel": "Reason for OASiS Status Change",
                 "Eligibility completed": "Eligibility Completed",
                 "Approver email": "Approver Email"
                 }

recrep_column_dict = {"APPROVAL_FLOW_STARTED": "Date Sent to AAS",
                      "OC_MEMBERSHIP_STATE": "OC Membership State",
                      "ARTICLETYPE": "Article Type",
                      "AUTHOREMAIL": "Author Email",
                      "PRS ARTICLE ID": "PRS Article ID",
                      "Unique Article ID": "Unique Article ID",
                      "APPROVED DATE": "Date Approved",
                      "REJECTED DATE": "Date Rejected",
                      "FORWARDED": "Date Forwarded",
                      "APPROVING_INSTITUTION": "Approving Institution",
                      "APPROVING_MANAGER": "Approver Email",
                      "ARTICLETITLE": "Article Title",
                      "JOURNAL NAME": "Journal Title",
                      "ARTICLE_MAPPING": "Article Type Assigned",
                      "LICENSETYPE": "Journal License",
                      "SUBMISSION_INSTITUTION": "Submission Institution",
                      "SUBMISSION_INSTITUTION_BPID": "Submission Institution BPID",
                      "PRS_INSTITUTION_TEXT": "PRS Institution",
                      "SELECTED_AFFILIATION_BPID": "Selected Affiliation BPID",
                      "IOCM_BY_IP_FOUND_INSTITUTION": "IP Institution",
                      "OCM_BY_IP_FOUND": "IP BPID",
                      "OCM_BY_EMAIL_FOUND_INSTITUTION": "Email Institution",
                      "OCM_BY_EMAIL_FOUND": "Email BPID",
                      "INITIAL_ASSIGNED_INSTITUTION": "Institution Assigned at Acceptance",
                      "INITIAL_ASSIGNED_INSTITUTION_BP": "Institution Assigned at Acceptance BPID",
                      "WAIVER_TYPE": "Waiver Type",
                      "CONTRACT_TYPE": "Contract Type",
                      "FIRSTNAME": "Author First Name",
                      "LASTNAME": "Author Last Name"
                      }

jflux_dict = {"articledoi_XLS": "DOI",
              "Agreement_XLS": "Agreement",
              "Journal ID_XLS": "Journal ID",
              "Monthly Reporting Status_XLS": "Monthly Reporting Status"
              }

email_country_dict = {"se": "BIBSAM",
                      "uk": "JISC",
                      "no": "SIKT",
                      "fi": "FINELIB",
                      "ie": "IREL",
                      "nl": "UKB",
                      "hu": "HAS",
                      "it": "CRUI-CARE",
                      "pl": "ICM",
                      "au": "CAUL",
                      "ca": "FSNL",
                      "at": "KEMÖ",
                      "ch": "CSAL",
                      "de": "DEAL",
                      "nz": "CAUL",
                      "qa": "QNL",
                      "es": "CRUE-CSIC"
                      }

articletype_dict = {"Original paper": "ORIGINALPAPER",
                    "Review paper": "REVIEWPAPER",
                    "Brief communication": "BRIEFCOMMUNICATION",
                    "Continuing education": "CONTINUINGEDUCATION",
                    "Book review": "BOOKREVIEW",
                    "Editorial notes": "EDITORIALNOTES",
                    "Report": "REPORT",
                    "Letter": "LETTER"}

# dicts for error trends

month_name = {1: "Jan",
              2: "Feb",
              3: "Mar",
              4: "Apr",
              5: "May",
              6: "Jun",
              7: "Jul",
              8: "Aug",
              9: "Sep",
              10: "Oct",
              11: "Nov",
              12: "Dec"}

doi_dict_1 = {"DOI": "Total DOIs"}
doi_dict_2 = {"DOI": "Yes/No"}
doi_dict_3 = {"Yes/No": "monthly count"}
doi_dict_4 = {"DOI": "count"}
doi_dict_5 = {"% as number": "percentage"}
doi_dict_6 = {"DOI": "count"}

error_trends_dict_2 = {"No": "Yes", 100: 0, "100.0%": "0%"}
error_trends_dict_3 = {"DOI": "Total DOIs",
                       "Verification Status (old DOI)": "Old verification status",
                       "Verification Status (new DOI)": "New verification status"}
