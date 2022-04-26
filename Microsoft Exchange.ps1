#
# Microsoft Exchange.ps1 - IDM System PowerShell Script for Microsoft Exchange Services.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


$Log_MaskableKeys = @(
    'password'
)


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'server'
                type = 'textbox'
                label = 'Server'
                description = 'Name of Exchange server'
                value = 'ps.outlook.com'
            }
            @{
                name = 'use_secure_connection'
                type = 'checkbox'
                label = 'Use secure connection'
                description = 'Using http or https'
                value = $false
            }
            @{
                name = 'skip_certificate_checks'
                type = 'checkgroup'
                label = 'Skip certificate checks'
                label_indent = $true
                description = 'Skip the following certificate checks'
                table = @{
                    rows = @(
                        @{ id = 'SkipCACheck';         display_text = 'Certificate signed by trusted certification authority (CA)' }
                        @{ id = 'SkipCNCheck';         display_text = 'Matching certificate common name (CN)' }
                        @{ id = 'SkipRevocationCheck'; display_text = 'Valid certificate revocation status' }
                    )
                    settings_checkgroup = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = @()
                hidden = '!use_secure_connection'
            }
            @{
                name = 'use_proxy_server'
                type = 'checkbox'
                label = 'Use proxy server'
                description = 'Behind a proxy server?'
                value = $false
            }
            @{
                name = 'proxy_access_type'
                type = 'combo'
                label = 'Proxy access type'
                label_indent = $true
                description = 'Proxy access type'
                table = @{
                    rows = @(
                        @{ id = 'AutoDetect';    display_text = 'AutoDetect' }
                        @{ id = 'IEConfig';      display_text = 'IEConfig' }
                        @{ id = 'WinHttpConfig'; display_text = 'WinHttpConfig' }
                    )
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = 'AutoDetect'
                hidden = '!use_proxy_server'
            }
            @{
                name = 'authentication'
                type = 'combo'
                label = 'Authentication'
                description = 'Authentication method'
                table = @{
                    rows = @(
                        @{ id = 'Default';                         display_text = 'Default' }
                        @{ id = 'Basic';                           display_text = 'Basic' }
                        @{ id = 'Credssp';                         display_text = 'Credssp' }
                        @{ id = 'Digest';                          display_text = 'Digest' }
                        @{ id = 'Kerberos';                        display_text = 'Kerberos' }
                        @{ id = 'Negotiate';                       display_text = 'Negotiate' }
                        @{ id = 'NegotiateWithImplicitCredential'; display_text = 'NegotiateWithImplicitCredential' }
                    )
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = 'Default'
            }
            @{
                name = 'use_svc_account_creds'
                type = 'checkbox'
                label = 'Use credentials of service account'
                value = $true
            }
            @{
                name = 'username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                description = 'User account name to access exchange'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                description = 'User account password to access exchange'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 10
            }
        )
    }

    if ($TestConnection) {
        Open-MsExchangeSession (ConvertFrom-Json2 $ConnectionParams)
    }

    if ($Configuration) {
        Open-MsExchangeSession (ConvertFrom-Json2 $ConnectionParams)

        @(
            @{
                name = 'organizational_unit'
                type = 'combo'
                label = 'Organizational unit'
                description = 'Organization Unit to start searching on; empty or * searches all'
                table = @{
                    rows = @( @{ display = '*'; value = '*' } ) + @( Get-MsExchangeOrganizationalUnit | Sort-Object -Property 'canonicalName' | ForEach-Object { @{ display = $_.canonicalName; value = $_.distinguishedName } } )
                    settings_combo = @{
                        display_column = 'display'
                        value_column = 'value'
                    }
                }
                value = '*'
            }
            @{
                name = 'domain_controller'
                type = 'textbox'
                label = 'Domain controller'
                description = 'Name of Domain Controller to target'
                value = ''
            }
        )
    }

    Log info "Done"
}


function Idm-OnUnload {
    Close-MsExchangeSession
}


#
# CRUD functions
#

$Properties = @{
    CASMailbox = @(
        @{ name = 'ActiveSyncAllowedDeviceIDs';                          set = $true;                  }
        @{ name = 'ActiveSyncBlockedDeviceIDs';                          set = $true;                  }
        @{ name = 'ActiveSyncDebugLogging';                              set = $true;                  }
        @{ name = 'ActiveSyncEnabled';                                   default = $true; set = $true; }
        @{ name = 'ActiveSyncMailboxPolicy';                             set = $true;                  }
        @{ name = 'ActiveSyncMailboxPolicyIsDefaulted';                                                }
        @{ name = 'ActiveSyncSuppressReadReceipt';                       set = $true;                  }
        @{ name = 'DisplayName';                                         set = $true;                  }
        @{ name = 'DistinguishedName';                                                                 }
        @{ name = 'ECPEnabled';                                          default = $true; set = $true; }
        @{ name = 'EmailAddresses';                                      set = $true;                  }
        @{ name = 'EwsAllowEntourage';                                   set = $true;                  }
        @{ name = 'EwsAllowList';                                        set = $true;                  }
        @{ name = 'EwsAllowMacOutlook';                                  set = $true;                  }
        @{ name = 'EwsAllowOutlook';                                     set = $true;                  }
        @{ name = 'EwsApplicationAccessPolicy';                          set = $true;                  }
        @{ name = 'EwsBlockList';                                        set = $true;                  }
        @{ name = 'EwsEnabled';                                          set = $true;                  }
        @{ name = 'ExchangeVersion';                                                                   }
        @{ name = 'ExternalImapSettings';                                                              }
        @{ name = 'ExternalPopSettings';                                                               }
        @{ name = 'ExternalSmtpSettings';                                                              }
        @{ name = 'Guid';                                                default = $true; key = $true; }
        @{ name = 'HasActiveSyncDevicePartnership';                      set = $true;                  }
        @{ name = 'Id';                                                  default = $true;              }
        @{ name = 'Identity';                                            default = $true;              }
        @{ name = 'ImapEnabled';                                         default = $true; set = $true; }
        @{ name = 'ImapEnableExactRFC822Size';                           set = $true;                  }
        @{ name = 'ImapForceICalForCalendarRetrievalOption';             set = $true;                  }
        @{ name = 'ImapMessagesRetrievalMimeFormat';                     set = $true;                  }
        @{ name = 'ImapSuppressReadReceipt';                             set = $true;                  }
        @{ name = 'ImapUseProtocolDefaults';                             set = $true;                  }
        @{ name = 'InternalImapSettings';                                                              }
        @{ name = 'InternalPopSettings';                                                               }
        @{ name = 'InternalSmtpSettings';                                                              }
        @{ name = 'IsOptimizedForAccessibility';                         set = $true;                  }
        @{ name = 'IsValid';                                             default = $true;              }
        @{ name = 'LegacyExchangeDN';                                                                  }
        @{ name = 'LinkedMasterAccount';                                 default = $true;              }
        @{ name = 'MAPIBlockOutlookExternalConnectivity';                set = $true;                  }
        @{ name = 'MAPIBlockOutlookNonCachedMode';                       set = $true;                  }
        @{ name = 'MAPIBlockOutlookRpcHttp';                             set = $true;                  }
        @{ name = 'MAPIBlockOutlookVersions';                            set = $true;                  }
        @{ name = 'MAPIEnabled';                                         set = $true;                  }
        @{ name = 'MapiHttpEnabled';                                     set = $true;                  }
        @{ name = 'Name';                                                set = $true;                  }
        @{ name = 'ObjectCategory';                                                                    }
        @{ name = 'ObjectClass';                                                                       }
        @{ name = 'ObjectState';                                                                       }
        @{ name = 'OrganizationId';                                                                    }
        @{ name = 'OriginatingServer';                                                                 }
        @{ name = 'OWAEnabled';                                          default = $true; set = $true; }
        @{ name = 'OWAforDevicesEnabled';                                set = $true;                  }
        @{ name = 'OwaMailboxPolicy';                                    set = $true;                  }
        @{ name = 'PopEnabled';                                          default = $true; set = $true; }
        @{ name = 'PopEnableExactRFC822Size';                            set = $true;                  }
        @{ name = 'PopForceICalForCalendarRetrievalOption';              set = $true;                  }
        @{ name = 'PopMessageDeleteEnabled';                                                           }
        @{ name = 'PopMessagesRetrievalMimeFormat';                      set = $true;                  }
        @{ name = 'PopSuppressReadReceipt';                              set = $true;                  }
        @{ name = 'PopUseProtocolDefaults';                              set = $true;                  }
        @{ name = 'PrimarySmtpAddress';                                  default = $true; set = $true; }
        @{ name = 'PSComputerName';                                                                    }
        @{ name = 'PSShowComputerName';                                                                }
        @{ name = 'PublicFolderClientAccess';                            set = $true;                  }
        @{ name = 'RunspaceId';                                                                        }
        @{ name = 'SamAccountName';                                      set = $true;                  }
        @{ name = 'ServerLegacyDN';                                                                    }
        @{ name = 'ServerName';                                                                        }
        @{ name = 'ShowGalAsDefaultView';                                set = $true;                  }
        @{ name = 'UniversalOutlookEnabled';                             set = $true;                  }
        @{ name = 'WhenChanged';                                                                       }
        @{ name = 'WhenChangedUTC';                                                                    }
        @{ name = 'WhenCreated';                                                                       }
        @{ name = 'WhenCreatedUTC';                                                                    }
    )

    Mailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                              set = $true;                                  }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                     set = $true;                                  }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';              set = $true;                                  }
        @{ name = 'AccountDisabled';                                     set = $true;                                  }
        @{ name = 'AddressBookPolicy';                                   enable = $true; set = $true;                  }
        @{ name = 'AddressListMembership';                                                                             }
        @{ name = 'AdminDisplayVersion';                                                                               }
        @{ name = 'AdministrativeUnits';                                                                               }
        @{ name = 'AggregatedMailboxGuids';                                                                            }
        @{ name = 'Alias';                                               default = $true; enable = $true; set = $true; }
        @{ name = 'AntispamBypassEnabled';                               set = $true;                                  }
        @{ name = 'ArbitrationMailbox';                                  set = $true;                                  }
        @{ name = 'ArchiveDatabase';                                     enable = $true; set = $true;                  }
        @{ name = 'ArchiveDomain';                                       enable = $true; set = $true;                  }
        @{ name = 'ArchiveGuid';                                         enable = $true;                               }
        @{ name = 'ArchiveName';                                         default = $true; enable = $true; set = $true; }
        @{ name = 'ArchiveQuota';                                        set = $true;                                  }
        @{ name = 'ArchiveRelease';                                                                                    }
        @{ name = 'ArchiveState';                                        set = $true;                                  }
        @{ name = 'ArchiveStatus';                                                                                     }
        @{ name = 'ArchiveWarningQuota';                                 set = $true;                                  }
        @{ name = 'AuditAdmin';                                          set = $true;                                  }
        @{ name = 'AuditDelegate';                                       set = $true;                                  }
        @{ name = 'AuditEnabled';                                        set = $true;                                  }
        @{ name = 'AuditLogAgeLimit';                                    set = $true;                                  }
        @{ name = 'AuditOwner';                                          set = $true;                                  }
        @{ name = 'AutoExpandingArchiveEnabled';                                                                       }
        @{ name = 'BypassModerationFromSendersOrMembers';                set = $true;                                  }
        @{ name = 'CalendarLoggingQuota';                                set = $true;                                  }
        @{ name = 'CalendarRepairDisabled';                              set = $true;                                  }
        @{ name = 'CalendarVersionStoreDisabled';                        set = $true;                                  }
        @{ name = 'ComplianceTagHoldApplied';                                                                          }
        @{ name = 'CustomAttribute1';                                    set = $true;                                  }
        @{ name = 'CustomAttribute2';                                    set = $true;                                  }
        @{ name = 'CustomAttribute3';                                    set = $true;                                  }
        @{ name = 'CustomAttribute4';                                    set = $true;                                  }
        @{ name = 'CustomAttribute5';                                    set = $true;                                  }
        @{ name = 'CustomAttribute6';                                    set = $true;                                  }
        @{ name = 'CustomAttribute7';                                    set = $true;                                  }
        @{ name = 'CustomAttribute8';                                    set = $true;                                  }
        @{ name = 'CustomAttribute9';                                    set = $true;                                  }
        @{ name = 'CustomAttribute10';                                   set = $true;                                  }
        @{ name = 'CustomAttribute11';                                   set = $true;                                  }
        @{ name = 'CustomAttribute12';                                   set = $true;                                  }
        @{ name = 'CustomAttribute13';                                   set = $true;                                  }
        @{ name = 'CustomAttribute14';                                   set = $true;                                  }
        @{ name = 'CustomAttribute15';                                   set = $true;                                  }
        @{ name = 'Database';                                            default = $true; enable = $true; set = $true; }
        @{ name = 'DataEncryptionPolicy';                                set = $true;                                  }
        @{ name = 'DefaultPublicFolderMailbox';                          set = $true;                                  }
        @{ name = 'DelayHoldApplied';                                                                                  }
        @{ name = 'DeliverToMailboxAndForward';                          set = $true;                                  }
        @{ name = 'DisabledArchiveDatabase';                                                                           }
        @{ name = 'DisabledArchiveGuid';                                                                               }
        @{ name = 'DisabledMailboxLocations';                                                                          }
        @{ name = 'DisableThrottling';                                   set = $true;                                  }
        @{ name = 'DisplayName';                                         default = $true; enable = $true; set = $true; }
        @{ name = 'DistinguishedName';                                                                                 }
        @{ name = 'DowngradeHighPriorityMessagesEnabled';                set = $true;                                  }
        @{ name = 'EffectivePublicFolderMailbox';                                                                      }
        @{ name = 'ElcProcessingDisabled';                               set = $true;                                  }
        @{ name = 'EmailAddresses';                                      default = $true; set = $true;                 }
        @{ name = 'EmailAddressPolicyEnabled';                           set = $true;                                  }
        @{ name = 'EndDateForRetentionHold';                             set = $true;                                  }
        @{ name = 'ExchangeGuid';                                                                                      }
        @{ name = 'ExchangeUserAccountControl';                                                                        }
        @{ name = 'ExchangeSecurityDescriptor';                                                                        }
        @{ name = 'ExchangeVersion';                                                                                   }
        @{ name = 'ExtensionCustomAttribute1';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute2';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute3';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute4';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute5';                           set = $true;                                  }
        @{ name = 'Extensions';                                                                                        }
        @{ name = 'ExternalDirectoryObjectId';                                                                         }
        @{ name = 'ExternalOofOptions';                                  set = $true;                                  }
        @{ name = 'ForwardingAddress';                                   set = $true;                                  }
        @{ name = 'ForwardingSmtpAddress';                               set = $true;                                  }
        @{ name = 'GeneratedOfflineAddressBooks';                                                                      }
        @{ name = 'GrantSendOnBehalfTo';                                 set = $true;                                  }
        @{ name = 'Guid';                                                default = $true; key = $true;                 }
        @{ name = 'HasPicture';                                                                                        }
        @{ name = 'HasSnackyAppData';                                                                                  }
        @{ name = 'HasSpokenName';                                                                                     }
        @{ name = 'HiddenFromAddressListsEnabled';                       set = $true;                                  }
        @{ name = 'Id';                                                  default = $true;                              }
        @{ name = 'Identity';                                                                                          }
        @{ name = 'ImListMigrationCompleted';                            set = $true;                                  }
        @{ name = 'ImmutableId';                                         set = $true;                                  }
        @{ name = 'InactiveMailboxRetireTime';                                                                         }
        @{ name = 'IncludeInGarbageCollection';                                                                        }
        @{ name = 'InPlaceHolds';                                                                                      }
        @{ name = 'IsDirSynced';                                                                                       }
        @{ name = 'IsExcludedFromServingHierarchy';                      set = $true;                                  }
        @{ name = 'IsHierarchyReady';                                    set = $true;                                  }
        @{ name = 'IsHierarchySyncEnabled';                              set = $true;                                  }
        @{ name = 'IsInactiveMailbox';                                                                                 }
        @{ name = 'IsLinked';                                                                                          }
        @{ name = 'IsMachineToPersonTextMessagingEnabled';                                                             }
        @{ name = 'IsMailboxEnabled';                                                                                  }
        @{ name = 'IsPersonToPersonTextMessagingEnabled';                                                              }
        @{ name = 'IsResource';                                                                                        }
        @{ name = 'IsRootPublicFolderMailbox';                                                                         }
        @{ name = 'IsShared';                                                                                          }
        @{ name = 'IsSoftDeletedByDisable';                                                                            }
        @{ name = 'IsSoftDeletedByRemove';                                                                             }
        @{ name = 'IssueWarningQuota';                                   set = $true;                                  }
        @{ name = 'IsValid';                                                                                           }
        @{ name = 'JournalArchiveAddress';                               set = $true;                                  }
        @{ name = 'Languages';                                           set = $true;                                  }
        @{ name = 'LastExchangeChangedTime';                                                                           }
        @{ name = 'LegacyExchangeDN';                                                                                  }
        @{ name = 'LinkedMasterAccount';                                 enable = $true; set = $true;                  }
        @{ name = 'LitigationHoldDate';                                  set = $true;                                  }
        @{ name = 'LitigationHoldDuration';                              set = $true;                                  }
        @{ name = 'LitigationHoldEnabled';                               set = $true;                                  }
        @{ name = 'LitigationHoldOwner';                                 set = $true;                                  }
        @{ name = 'MailboxContainerGuid';                                                                              }
        @{ name = 'MailboxLocations';                                                                                  }
        @{ name = 'MailboxMoveBatchName';                                                                              }
        @{ name = 'MailboxMoveFlags';                                                                                  }
        @{ name = 'MailboxMoveRemoteHostName';                                                                         }
        @{ name = 'MailboxMoveSourceMDB';                                                                              }
        @{ name = 'MailboxMoveStatus';                                                                                 }
        @{ name = 'MailboxMoveTargetMDB';                                                                              }
        @{ name = 'MailboxPlan';                                                                                       }
        @{ name = 'MailboxProvisioningConstraint';                                                                     }
        @{ name = 'MailboxProvisioningPreferences';                                                                    }
        @{ name = 'MailboxRegion';                                       set = $true;                                  }
        @{ name = 'MailboxRegionLastUpdateTime';                                                                       }
        @{ name = 'MailboxRelease';                                                                                    }
        @{ name = 'MailTip';                                             set = $true;                                  }
        @{ name = 'MailTipTranslations';                                 set = $true;                                  }
        @{ name = 'ManagedFolderMailboxPolicy';                          enable = $true; set = $true;                  }
        @{ name = 'MaxBlockedSenders';                                   set = $true;                                  }
        @{ name = 'MaxReceiveSize';                                      set = $true;                                  }
        @{ name = 'MaxSafeSenders';                                      set = $true;                                  }
        @{ name = 'MaxSendSize';                                         set = $true;                                  }
        @{ name = 'MessageCopyForSendOnBehalfEnabled';                   set = $true;                                  }
        @{ name = 'MessageCopyForSentAsEnabled';                         set = $true;                                  }
        @{ name = 'MessageTrackingReadStatusEnabled';                                                                  }
        @{ name = 'MicrosoftOnlineServicesID';                           set = $true;                                  }
        @{ name = 'ModeratedBy';                                         set = $true;                                  }
        @{ name = 'ModerationEnabled';                                   set = $true;                                  }
        @{ name = 'Name';                                                set = $true;                                  }
        @{ name = 'NetID';                                                                                             }
        @{ name = 'ObjectCategory';                                                                                    }
        @{ name = 'ObjectClass';                                                                                       }
        @{ name = 'ObjectState';                                                                                       }
        @{ name = 'Office';                                              set = $true;                                  }
        @{ name = 'OfflineAddressBook';                                  set = $true;                                  }
        @{ name = 'OrganizationalUnit';                                                                                }
        @{ name = 'OrganizationId';                                                                                    }
        @{ name = 'OriginatingServer';                                                                                 }
        @{ name = 'OrphanSoftDeleteTrackingTime';                                                                      }
        @{ name = 'PersistedCapabilities';                                                                             }
        @{ name = 'PoliciesExcluded';                                                                                  }
        @{ name = 'PoliciesIncluded';                                                                                  }
        @{ name = 'PrimarySmtpAddress';                                  default = $true; enable = $true; set = $true; }
        @{ name = 'ProhibitSendQuota';                                   set = $true;                                  }
        @{ name = 'ProhibitSendReceiveQuota';                            set = $true;                                  }
        @{ name = 'ProtocolSettings';                                                                                  }
        @{ name = 'PSComputerName';                                                                                    }
        @{ name = 'PSShowComputerName';                                                                                }
        @{ name = 'QueryBaseDN';                                         set = $true;                                  }
        @{ name = 'QueryBaseDNRestrictionEnabled';                                                                     }
        @{ name = 'RecipientLimits';                                     set = $true;                                  }
        @{ name = 'RecipientType';                                                                                     }
        @{ name = 'RecipientTypeDetails';                                                                              }
        @{ name = 'ReconciliationId';                                                                                  }
        @{ name = 'RecoverableItemsQuota';                               set = $true;                                  }
        @{ name = 'RecoverableItemsWarningQuota';                        set = $true;                                  }
        @{ name = 'RejectMessagesFrom';                                  set = $true;                                  }
        @{ name = 'RejectMessagesFromDLMembers';                         set = $true;                                  }
        @{ name = 'RejectMessagesFromSendersOrMembers';                  set = $true;                                  }
        @{ name = 'RemoteAccountPolicy';                                                                               }
        @{ name = 'RemoteRecipientType';                                 set = $true;                                  }
        @{ name = 'RequireSenderAuthenticationEnabled';                  set = $true;                                  }
        @{ name = 'ResetPasswordOnNextLogon';                            set = $true;                                  }
        @{ name = 'ResourceCapacity';                                    set = $true;                                  }
        @{ name = 'ResourceCustom';                                      set = $true;                                  }
        @{ name = 'ResourceType';                                                                                      }
        @{ name = 'RetainDeletedItemsFor';                               set = $true;                                  }
        @{ name = 'RetainDeletedItemsUntilBackup';                       set = $true;                                  }
        @{ name = 'RetentionComment';                                    set = $true;                                  }
        @{ name = 'RetentionHoldEnabled';                                set = $true;                                  }
        @{ name = 'RetentionPolicy';                                     enable = $true; set = $true;                  }
        @{ name = 'RetentionUrl';                                        set = $true;                                  }
        @{ name = 'RoleAssignmentPolicy';                                enable = $true; set = $true;                  }
        @{ name = 'RoomMailboxAccountEnabled';                                                                         }
        @{ name = 'RulesQuota';                                          set = $true;                                  }
        @{ name = 'RunspaceId';                                                                                        }
        @{ name = 'SamAccountName';                                      set = $true;                                  }
        @{ name = 'SCLDeleteEnabled';                                    set = $true;                                  }
        @{ name = 'SCLDeleteThreshold';                                  set = $true;                                  }
        @{ name = 'SCLJunkEnabled';                                      set = $true;                                  }
        @{ name = 'SCLJunkThreshold';                                    set = $true;                                  }
        @{ name = 'SCLQuarantineEnabled';                                set = $true;                                  }
        @{ name = 'SCLQuarantineThreshold';                              set = $true;                                  }
        @{ name = 'SCLRejectEnabled';                                    set = $true;                                  }
        @{ name = 'SCLRejectThreshold';                                  set = $true;                                  }
        @{ name = 'SendModerationNotifications';                         set = $true;                                  }
        @{ name = 'ServerLegacyDN';                                                                                    }
        @{ name = 'ServerName';                                                                                        }
        @{ name = 'SharingPolicy';                                       set = $true;                                  }
        @{ name = 'SiloName';                                                                                          }
        @{ name = 'SimpleDisplayName';                                   set = $true;                                  }
        @{ name = 'SingleItemRecoveryEnabled';                           set = $true;                                  }
        @{ name = 'SKUAssigned';                                                                                       }
        @{ name = 'SourceAnchor';                                                                                      }
        @{ name = 'StartDateForRetentionHold';                           set = $true;                                  }
        @{ name = 'StsRefreshTokensValidFrom';                           set = $true;                                  }
        @{ name = 'ThrottlingPolicy';                                    set = $true;                                  }
        @{ name = 'UMDtmfMap';                                           set = $true;                                  }
        @{ name = 'UMEnabled';                                                                                         }
        @{ name = 'UnifiedMailbox';                                                                                    }
        @{ name = 'UsageLocation';                                                                                     }
        @{ name = 'UseDatabaseQuotaDefaults';                            set = $true;                                  }
        @{ name = 'UseDatabaseRetentionDefaults';                        set = $true;                                  }
        @{ name = 'UserCertificate';                                     set = $true;                                  }
        @{ name = 'UserPrincipalName';                                   set = $true;                                  }
        @{ name = 'UserSMimeCertificate';                                set = $true;                                  }
        @{ name = 'WasInactiveMailbox';                                                                                }
        @{ name = 'WhenChanged';                                                                                       }
        @{ name = 'WhenChangedUTC';                                                                                    }
        @{ name = 'WhenCreated';                                                                                       }
        @{ name = 'WhenCreatedUTC';                                                                                    }
        @{ name = 'WhenMailboxCreated';                                                                                }
        @{ name = 'WhenSoftDeleted';                                                                                   }
        @{ name = 'WindowsEmailAddress';                                 set = $true;                                  }
        @{ name = 'WindowsLiveID';                                                                                     }
    )

    MailboxDatabase = @(
        @{ name = 'ActivationPreference';                                                              }
        @{ name = 'AdminDisplayName';                                                                  }
        @{ name = 'AdminDisplayVersion';                                                               }
        @{ name = 'AdministrativeGroup';                                                               }
        @{ name = 'AllDatabaseCopies';                                                                 }
        @{ name = 'AllowFileRestore';                                                                  }
        @{ name = 'AutoDagExcludeFromMonitoring';                                                      }
        @{ name = 'AutoDatabaseMountDial';                                                             }
        @{ name = 'AvailableNewMailboxSpace';                                                          }
        @{ name = 'BackgroundDatabaseMaintenance';                                                     }
        @{ name = 'BackgroundDatabaseMaintenanceDelay';                                                }
        @{ name = 'BackgroundDatabaseMaintenanceSerialization';                                        }
        @{ name = 'BackupInProgress';                                                                  }
        @{ name = 'CachedClosedTables';                                                                }
        @{ name = 'CachePriority';                                                                     }
        @{ name = 'CafeEndpoints';                                                                     }
        @{ name = 'CalendarLoggingQuota';                                                              }
        @{ name = 'CircularLoggingEnabled';                                                            }
        @{ name = 'CreationSchemaVersion';                                                             }
        @{ name = 'CurrentSchemaVersion';                                                              }
        @{ name = 'DatabaseCopies';                                                                    }
        @{ name = 'DatabaseCreated';                                                                   }
        @{ name = 'DatabaseExtensionSize';                                                             }
        @{ name = 'DatabaseGroup';                                                                     }
        @{ name = 'DatabaseSize';                                        default = $true;              }
        @{ name = 'DataMoveReplicationConstraint';                                                     }
        @{ name = 'DeletedItemRetention';                                                              }
        @{ name = 'Description';                                         default = $true;              }
        @{ name = 'DistinguishedName';                                                                 }
        @{ name = 'DumpsterServersNotAvailable';                                                       }
        @{ name = 'DumpsterStatistics';                                                                }
        @{ name = 'EdbFilePath';                                                                       }
        @{ name = 'EventHistoryRetentionPeriod';                                                       }
        @{ name = 'ExchangeLegacyDN';                                                                  }
        @{ name = 'ExchangeVersion';                                     default = $true;              }
        @{ name = 'Guid';                                                                              }
        @{ name = 'Id';                                                  default = $true;              }
        @{ name = 'Identity';                                            default = $true; key = $true; }
        @{ name = 'IndexEnabled';                                                                      }
        @{ name = 'InvalidDatabaseCopies';                                                             }
        @{ name = 'IsExcludedFromInitialProvisioning';                                                 }
        @{ name = 'IsExcludedFromProvisioning';                                                        }
        @{ name = 'IsExcludedFromProvisioningBy';                                                      }
        @{ name = 'IsExcludedFromProvisioningByOperator';                                              }
        @{ name = 'IsExcludedFromProvisioningBySchemaVersionMonitoring';                               }
        @{ name = 'IsExcludedFromProvisioningBySpaceMonitoring';                                       }
        @{ name = 'IsExcludedFromProvisioningDueToLogicalCorruption';                                  }
        @{ name = 'IsExcludedFromProvisioningForDraining';                                             }
        @{ name = 'IsExcludedFromProvisioningReason';                                                  }
        @{ name = 'IsMailboxDatabase';                                   default = $true;              }
        @{ name = 'IsPublicFolderDatabase';                              default = $true;              }
        @{ name = 'IssueWarningQuota';                                   default = $true;              }
        @{ name = 'IsSuspendedFromProvisioning';                         default = $true;              }
        @{ name = 'IsValid';                                             default = $true;              }
        @{ name = 'JournalRecipient';                                                                  }
        @{ name = 'LastCopyBackup';                                                                    }
        @{ name = 'LastDifferentialBackup';                                                            }
        @{ name = 'LastFullBackup';                                                                    }
        @{ name = 'LastIncrementalBackup';                                                             }
        @{ name = 'LogBuffers';                                                                        }
        @{ name = 'LogCheckpointDepth';                                                                }
        @{ name = 'LogFilePrefix';                                                                     }
        @{ name = 'LogFileSize';                                                                       }
        @{ name = 'LogFolderPath';                                                                     }
        @{ name = 'MailboxLoadBalanceEnabled';                                                         }
        @{ name = 'MailboxLoadBalanceMaximumEdbFileSize';                                              }
        @{ name = 'MailboxLoadBalanceOverloadedThreshold';                                             }
        @{ name = 'MailboxLoadBalanceRelativeLoadCapacity';                                            }
        @{ name = 'MailboxLoadBalanceUnderloadedThreshold';                                            }
        @{ name = 'MailboxProvisioningAttributes';                                                     }
        @{ name = 'MailboxRetention';                                                                  }
        @{ name = 'MaintenanceSchedule';                                                               }
        @{ name = 'MasterServerOrAvailabilityGroup';                                                   }
        @{ name = 'MasterType';                                                                        }
        @{ name = 'MaximumBackgroundDatabaseMaintenanceInterval';                                      }
        @{ name = 'MaximumCursors';                                                                    }
        @{ name = 'MaximumOpenTables';                                                                 }
        @{ name = 'MaximumPreReadPages';                                                               }
        @{ name = 'MaximumReplayPreReadPages';                                                         }
        @{ name = 'MaximumSessions';                                                                   }
        @{ name = 'MaximumTemporaryTables';                                                            }
        @{ name = 'MaximumVersionStorePages';                                                          }
        @{ name = 'MCDBAvailableSpace';                                                                }
        @{ name = 'MCDBSize';                                                                          }
        @{ name = 'MetaCacheDatabaseFilePath';                                                         }
        @{ name = 'MetaCacheDatabaseFolderPath';                                                       }
        @{ name = 'MetaCacheDatabaseMaxCapacityInBytes';                                               }
        @{ name = 'MetaCacheDatabaseMountpointFolderPath';                                             }
        @{ name = 'MetaCacheDatabaseRootFolderPath';                                                   }
        @{ name = 'MetaCacheDatabaseVolumesRootFolderPath';                                            }
        @{ name = 'MimimumBackgroundDatabaseMaintenanceInterval';                                      }
        @{ name = 'MountAtStartup';                                                                    }
        @{ name = 'Mounted';                                                                           }
        @{ name = 'MountedOnServer';                                                                   }
        @{ name = 'Name';                                                                              }
        @{ name = 'ObjectCategory';                                                                    }
        @{ name = 'ObjectClass';                                                                       }
        @{ name = 'ObjectState';                                                                       }
        @{ name = 'OfflineAddressBook';                                                                }
        @{ name = 'Organization';                                                                      }
        @{ name = 'OrganizationId';                                                                    }
        @{ name = 'OriginalDatabase';                                                                  }
        @{ name = 'OriginatingServer';                                   default = $true;              }
        @{ name = 'PreferredVersionStorePages';                                                        }
        @{ name = 'ProhibitSendQuota';                                                                 }
        @{ name = 'ProhibitSendReceiveQuota';                                                          }
        @{ name = 'PSComputerName';                                                                    }
        @{ name = 'PSShowComputerName';                                                                }
        @{ name = 'PublicFolderDatabase';                                                              }
        @{ name = 'QuotaNotificationSchedule';                                                         }
        @{ name = 'RecoverableItemsQuota';                                                             }
        @{ name = 'RecoverableItemsWarningQuota';                                                      }
        @{ name = 'Recovery';                                                                          }
        @{ name = 'ReplayBackgroundDatabaseMaintenance';                                               }
        @{ name = 'ReplayBackgroundDatabaseMaintenanceDelay';                                          }
        @{ name = 'ReplayCachePriority';                                                               }
        @{ name = 'ReplayCheckpointDepth';                                                             }
        @{ name = 'ReplayLagTimes';                                                                    }
        @{ name = 'ReplicationType';                                                                   }
        @{ name = 'RequestedSchemaVersion';                                                            }
        @{ name = 'RetainDeletedItemsUntilBackup';                                                     }
        @{ name = 'RpcClientAccessServer';                                                             }
        @{ name = 'RunspaceId';                                                                        }
        @{ name = 'Server';                                                                            }
        @{ name = 'ServerName';                                                                        }
        @{ name = 'Servers';                                                                           }
        @{ name = 'SnapshotLastCopyBackup';                                                            }
        @{ name = 'SnapshotLastDifferentialBackup';                                                    }
        @{ name = 'SnapshotLastFullBackup';                                                            }
        @{ name = 'SnapshotLastIncrementalBackup';                                                     }
        @{ name = 'TemporaryDataFolderPath';                                                           }
        @{ name = 'TruncationLagTimes';                                                                }
        @{ name = 'WhenChanged';                                                                       }
        @{ name = 'WhenChangedUTC';                                                                    }
        @{ name = 'WhenCreated';                                                                       }
        @{ name = 'WhenCreatedUTC';                                                                    }
        @{ name = 'WorkerProcessId';                                                                   }
    )

    MailboxStatistics = @(
        @{ name = 'AssociatedItemCount';                                                               }
        @{ name = 'AttachmentTableAvailableSize';                                                      }
        @{ name = 'AttachmentTableTotalSize';                                                          }
        @{ name = 'CurrentSchemaVersion';                                                              }
        @{ name = 'Database';                                                                          }
        @{ name = 'DatabaseName';                                        default = $true;              }
        @{ name = 'DatabaseIssueWarningQuota';                                                         }
        @{ name = 'DatabaseProhibitSendQuota';                                                         }
        @{ name = 'DatabaseProhibitSendReceiveQuota';                                                  }
        @{ name = 'DeletedItemCount';                                                                  }
        @{ name = 'DisconnectDate';                                      default = $true;              }
        @{ name = 'DisconnectReason';                                    default = $true;              }
        @{ name = 'DisplayName';                                         default = $true;              }
        @{ name = 'DumpsterMessagesPerFolderCountReceiveQuota';                                        }
        @{ name = 'DumpsterMessagesPerFolderCountWarningQuota';                                        }
        @{ name = 'ExternalDirectoryOrganizationId';                                                   }
        @{ name = 'FolderHierarchyChildrenCountReceiveQuota';                                          }
        @{ name = 'FolderHierarchyChildrenCountWarningQuota';                                          }
        @{ name = 'FolderHierarchyDepthReceiveQuota';                                                  }
        @{ name = 'FolderHierarchyDepthWarningQuota';                                                  }
        @{ name = 'FoldersCountReceiveQuota';                                                          }
        @{ name = 'FoldersCountWarningQuota';                                                          }
        @{ name = 'Identity';                                            default = $true; key = $true; }
        @{ name = 'IsArchiveMailbox';                                    default = $true;              }
        @{ name = 'IsDatabaseCopyActive';                                                              }
        @{ name = 'IsMoveDestination';                                                                 }
        @{ name = 'IsQuarantined';                                       default = $true;              }
        @{ name = 'IsValid';                                                                           }
        @{ name = 'ItemCount';                                                                         }
        @{ name = 'LastLoggedOnUserAccount';                             default = $true;              }
        @{ name = 'LastLogoffTime';                                                                    }
        @{ name = 'LastLogonTime';                                       default = $true;              }
        @{ name = 'LegacyDN';                                            default = $true;              }
        @{ name = 'MailboxGuid';                                         default = $true;              }
        @{ name = 'MailboxMessagesPerFolderCountReceiveQuota';                                         }
        @{ name = 'MailboxMessagesPerFolderCountWarningQuota';                                         }
        @{ name = 'MailboxTableIdentifier';                                                            }
        @{ name = 'MailboxType';                                                                       }
        @{ name = 'MapiIdentity';                                                                      }
        @{ name = 'MessageTableAvailableSize';                                                         }
        @{ name = 'MessageTableTotalSize';                                                             }
        @{ name = 'MoveHistory';                                                                       }
        @{ name = 'NamedPropertiesCountQuota';                                                         }
        @{ name = 'ObjectClass';                                                                       }
        @{ name = 'ObjectState';                                                                       }
        @{ name = 'OriginatingServer';                                                                 }
        @{ name = 'OtherTablesAvailableSize';                                                          }
        @{ name = 'OtherTablesTotalSize';                                                              }
        @{ name = 'PSComputerName';                                                                    }
        @{ name = 'PSShowComputerName';                                                                }
        @{ name = 'QuarantineDescription';                                                             }
        @{ name = 'QuarantineEnd';                                                                     }
        @{ name = 'QuarantineFileVersion';                                                             }
        @{ name = 'QuarantineLastCrash';                                                               }
        @{ name = 'RunspaceId';                                                                        }
        @{ name = 'ServerName';                                                                        }
        @{ name = 'StorageLimitStatus';                                                                }
        @{ name = 'TotalDeletedItemSize';                                                              }
        @{ name = 'TotalItemSize';                                       default = $true;              }
    )

    RemoteMailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                              set = $true;                                  }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                     set = $true;                                  }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';              set = $true;                                  }
        @{ name = 'AccountDisabled';                                                                                   }
        @{ name = 'AddressListMembership';                                                                             }
        @{ name = 'AggregatedMailboxGuids';                                                                            }
        @{ name = 'Alias';                                               default = $true; enable = $true; set = $true; }
        @{ name = 'ArbitrationMailbox';                                                                                }
        @{ name = 'ArchiveDatabase';                                                                                   }
        @{ name = 'ArchiveGuid';                                         set = $true;                                  }
        @{ name = 'ArchiveName';                                         default = $true; enable = $true; set = $true; }
        @{ name = 'ArchiveQuota';                                                                                      }
        @{ name = 'ArchiveRelease';                                                                                    }
        @{ name = 'ArchiveState';                                                                                      }
        @{ name = 'ArchiveStatus';                                                                                     }
        @{ name = 'ArchiveWarningQuota';                                                                               }
        @{ name = 'BypassModerationFromSendersOrMembers';                set = $true;                                  }
        @{ name = 'CalendarVersionStoreDisabled';                                                                      }
        @{ name = 'CustomAttribute1';                                    set = $true;                                  }
        @{ name = 'CustomAttribute2';                                    set = $true;                                  }
        @{ name = 'CustomAttribute3';                                    set = $true;                                  }
        @{ name = 'CustomAttribute4';                                    set = $true;                                  }
        @{ name = 'CustomAttribute5';                                    set = $true;                                  }
        @{ name = 'CustomAttribute6';                                    set = $true;                                  }
        @{ name = 'CustomAttribute7';                                    set = $true;                                  }
        @{ name = 'CustomAttribute8';                                    set = $true;                                  }
        @{ name = 'CustomAttribute9';                                    set = $true;                                  }
        @{ name = 'CustomAttribute10';                                   set = $true;                                  }
        @{ name = 'CustomAttribute11';                                   set = $true;                                  }
        @{ name = 'CustomAttribute12';                                   set = $true;                                  }
        @{ name = 'CustomAttribute13';                                   set = $true;                                  }
        @{ name = 'CustomAttribute14';                                   set = $true;                                  }
        @{ name = 'CustomAttribute15';                                   set = $true;                                  }
        @{ name = 'DeliverToMailboxAndForward';                                                                        }
        @{ name = 'DisabledArchiveDatabase';                                                                           }
        @{ name = 'DisabledArchiveGuid';                                                                               }
        @{ name = 'DisplayName';                                         enable = $true; set = $true;                  }
        @{ name = 'DistinguishedName';                                   default = $true;                              }
        @{ name = 'EmailAddresses';                                      set = $true;                                  }
        @{ name = 'EmailAddressPolicyEnabled';                           set = $true;                                  }
        @{ name = 'EndDateForRetentionHold';                                                                           }
        @{ name = 'ExchangeGuid';                                        set = $true;                                  }
        @{ name = 'ExchangeVersion';                                                                                   }
        @{ name = 'ExchangeUserAccountControl';                                                                        }
        @{ name = 'ExtensionCustomAttribute1';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute2';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute3';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute4';                           set = $true;                                  }
        @{ name = 'ExtensionCustomAttribute5';                           set = $true;                                  }
        @{ name = 'Extensions';                                                                                        }
        @{ name = 'ExternalDirectoryObjectId';                                                                         }
        @{ name = 'ForwardingAddress';                                                                                 }
        @{ name = 'GrantSendOnBehalfTo';                                 set = $true;                                  }
        @{ name = 'Guid';                                                default = $true; key = $true;                 }
        @{ name = 'HasPicture';                                                                                        }
        @{ name = 'HasSpokenName';                                                                                     }
        @{ name = 'HiddenFromAddressListsEnabled';                       set = $true;                                  }
        @{ name = 'Id';                                                  default = $true;                              }
        @{ name = 'Identity';                                                                                          }
        @{ name = 'ImmutableId';                                         set = $true;                                  }
        @{ name = 'InPlaceHolds';                                                                                      }
        @{ name = 'IsSoftDeletedByDisable';                                                                            }
        @{ name = 'IsSoftDeletedByRemove';                                                                             }
        @{ name = 'IsValid';                                                                                           }
        @{ name = 'JournalArchiveAddress';                                                                             }
        @{ name = 'LastExchangeChangedTime';                                                                           }
        @{ name = 'LegacyExchangeDN';                                                                                  }
        @{ name = 'LitigationHoldDate';                                                                                }
        @{ name = 'LitigationHoldEnabled';                                                                             }
        @{ name = 'LitigationHoldOwner';                                                                               }
        @{ name = 'MailboxContainerGuid';                                                                              }
        @{ name = 'MailboxLocations';                                                                                  }
        @{ name = 'MailboxMoveBatchName';                                                                              }
        @{ name = 'MailboxMoveFlags';                                                                                  }
        @{ name = 'MailboxMoveRemoteHostName';                                                                         }
        @{ name = 'MailboxMoveSourceMDB';                                                                              }
        @{ name = 'MailboxMoveStatus';                                                                                 }
        @{ name = 'MailboxMoveTargetMDB';                                                                              }
        @{ name = 'MailboxProvisioningConstraint';                                                                     }
        @{ name = 'MailboxProvisioningPreferences';                                                                    }
        @{ name = 'MailboxRelease';                                                                                    }
        @{ name = 'MailTip';                                             set = $true;                                  }
        @{ name = 'MailTipTranslations';                                 set = $true;                                  }
        @{ name = 'MaxReceiveSize';                                                                                    }
        @{ name = 'MaxSendSize';                                                                                       }
        @{ name = 'ModeratedBy';                                         set = $true;                                  }
        @{ name = 'ModerationEnabled';                                   set = $true;                                  }
        @{ name = 'Name';                                                default = $true; set = $true;                 }
        @{ name = 'ObjectCategory';                                                                                    }
        @{ name = 'ObjectClass';                                                                                       }
        @{ name = 'ObjectState';                                                                                       }
        @{ name = 'OnPremisesOrganizationalUnit';                                                                      }
        @{ name = 'OrganizationId';                                                                                    }
        @{ name = 'OriginatingServer';                                                                                 }
        @{ name = 'PersistedCapabilities';                                                                             }
        @{ name = 'PoliciesExcluded';                                                                                  }
        @{ name = 'PoliciesIncluded';                                                                                  }
        @{ name = 'ProtocolSettings';                                                                                  }
        @{ name = 'PrimarySmtpAddress';                                  default = $true; enable = $true; set = $true; }
        @{ name = 'PSComputerName';                                                                                    }
        @{ name = 'PSShowComputerName';                                                                                }
        @{ name = 'RecipientLimits';                                                                                   }
        @{ name = 'RecipientType';                                                                                     }
        @{ name = 'RecipientTypeDetails';                                                                              }
        @{ name = 'RecoverableItemsQuota';                               set = $true;                                  }
        @{ name = 'RecoverableItemsWarningQuota';                        set = $true;                                  }
        @{ name = 'RejectMessagesFrom';                                  set = $true;                                  }
        @{ name = 'RejectMessagesFromDLMembers';                         set = $true;                                  }
        @{ name = 'RejectMessagesFromSendersOrMembers';                  set = $true;                                  }
        @{ name = 'RemoteRecipientType';                                                                               }
        @{ name = 'RemoteRoutingAddress';                                default = $true; enable = $true; set = $true; }
        @{ name = 'RequireSenderAuthenticationEnabled';                  set = $true;                                  }
        @{ name = 'ResetPasswordOnNextLogon';                            set = $true;                                  }
        @{ name = 'RetainDeletedItemsFor';                                                                             }
        @{ name = 'RetentionComment';                                                                                  }
        @{ name = 'RetentionHoldEnabled';                                                                              }
        @{ name = 'RetentionUrl';                                                                                      }
        @{ name = 'RunspaceId';                                                                                        }
        @{ name = 'SamAccountName';                                      set = $true;                                  }
        @{ name = 'SendModerationNotifications';                         set = $true;                                  }
        @{ name = 'SimpleDisplayName';                                                                                 }
        @{ name = 'SingleItemRecoveryEnabled';                                                                         }
        @{ name = 'StartDateForRetentionHold';                                                                         }
        @{ name = 'StsRefreshTokensValidFrom';                                                                         }
        @{ name = 'UMDtmfMap';                                                                                         }
        @{ name = 'UserCertificate';                                                                                   }
        @{ name = 'UserPrincipalName';                                   set = $true;                                  }
        @{ name = 'UserSMimeCertificate';                                                                              }
        @{ name = 'WhenChanged';                                                                                       }
        @{ name = 'WhenChangedUTC';                                                                                    }
        @{ name = 'WhenCreated';                                                                                       }
        @{ name = 'WhenCreatedUTC';                                                                                    }
        @{ name = 'WhenMailboxCreated';                                                                                }
        @{ name = 'WhenSoftDeleted';                                                                                   }
        @{ name = 'WindowsEmailAddress';                                 set = $true;                                  }
    )
}


# Default properties and IDM properties are the same
foreach ($key in $Properties.Keys) {
    for ($i = 0; $i -lt $Properties.$key.Count; $i++) {
        if ($Properties.$key[$i].default) {
            $Properties.$key[$i].idm = $true
        }
    }
}


function Idm-CASMailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'CASMailbox'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'Unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.CASMailbox | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/client-access/get-casmailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeCASMailbox" -In @call_params
            Get-MsExchangeCASMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-CASMailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.CASMailbox | Where-Object { !$_.key -and !$_.set } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.CASMailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params
        
        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/client-access/set-casmailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeCASMailbox" -In @call_params
                $rv = Set-MsExchangeCASMailbox @call_params
            LogIO info "Set-MsExchangeCASMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxEnable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.enable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/enable-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Enable-MsExchangeMailbox" -In @call_params
                $rv = Enable-MsExchangeMailbox @call_params
            LogIO info "Enable-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'Mailbox'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'Unlimited'
        }

        if ($system_params.organizational_unit.length -gt 0 -and $system_params.organizational_unit -ne '*') {
            $call_params.OrganizationalUnit = $system_params.organizational_unit
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.Mailbox | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeMailbox" -In @call_params
            Get-MsExchangeMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.set } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Set-MsExchangeMailbox" -In @call_params
                $rv = Set-MsExchangeMailbox @call_params
            LogIO info "Set-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxDisable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.disable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/disable-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Disable-MsExchangeMailbox" -In @call_params
                $rv = Disable-MsExchangeMailbox @call_params
            LogIO info "Disable-MsExchangeMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxDatabasesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'MailboxDatabase'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{}

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.MailboxDatabase | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.MailboxDatabase | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailbox-databases-and-servers/get-mailboxdatabase?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # x Cloud

            LogIO info "Get-MsExchangeMailboxDatabase" -In
            Get-MsExchangeMailboxDatabase | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxPermissionAdd {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.permissionAdd } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Add-MsExchangeMailboxPermission" -In @call_params
                $rv = Add-MsExchangeMailboxPermission @call_params
            LogIO info "Add-MsExchangeMailboxPermission" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxPermissionRemove {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.key -and !$_.permissionRemove } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.Mailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/remove-mailboxpermission?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Remove-MsExchangeMailboxPermission" -In @call_params
                $rv = Remove-MsExchangeMailboxPermission @call_params
            LogIO info "Remove-MsExchangeMailboxPermission" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-MailboxStatisticsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'MailboxStatistics'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailboxstatistics?view=exchange-ps
            #
            # Parameter availability:
            # v On-premises
            # x Cloud

            Server = $system_params.server
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.MailboxStatistics | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.MailboxStatistics | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailboxstatistics?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeMailboxStatistics" -In @call_params
            Get-MsExchangeMailboxStatistics @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-RemoteMailboxEnable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.key -and !$_.enable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/federation-and-hybrid/enable-remotemailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # x Cloud

            LogIO info "Enable-MsExchangeRemoteMailbox" -In @call_params
                $rv = Enable-MsExchangeRemoteMailbox @call_params
            LogIO info "Enable-MsExchangeRemoteMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-RemoteMailboxesRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        Get-ClassMetaData -SystemParams $SystemParams -Class 'RemoteMailbox'
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $call_params = @{
            ResultSize = 'Unlimited'
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.RemoteMailbox | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/federation-and-hybrid/get-remotemailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # x Cloud

            LogIO info "Get-MsExchangeRemoteMailbox" -In @call_params
            Get-MsExchangeRemoteMailbox @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-RemoteMailboxSet {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.key -and !$_.set } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/federation-and-hybrid/set-remotemailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # x Cloud

            LogIO info "Set-MsExchangeRemoteMailbox" -In @call_params
                $rv = Set-MsExchangeRemoteMailbox @call_params
            LogIO info "Set-MsExchangeRemoteMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


function Idm-RemoteMailboxDisable {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.key -and !$_.disable } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

               #@{ name = '*'; allowance = 'optional' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        Open-MsExchangeSession $system_params

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.key }).name

        $call_params = @{
            Identity = $function_params.$key
            Confirm  = $false   # Be non-interactive
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/federation-and-hybrid/disable-remotemailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # x Cloud

            LogIO info "Disable-MsExchangeRemoteMailbox" -In @call_params
                $rv = Disable-MsExchangeRemoteMailbox @call_params
            LogIO info "Disable-MsExchangeRemoteMailbox" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}


#
# Helper functions
#

function Open-MsExchangeSession {
    param (
        [hashtable] $SystemParams
    )

    # Use connection related parameters only
    $connection_params = [ordered]@{
        server                  = $SystemParams.server
        use_secure_connection   = $SystemParams.use_secure_connection
        skip_certificate_checks = $SystemParams.skip_certificate_checks
        use_proxy_server        = $SystemParams.use_proxy_server
        proxy_access_type       = $SystemParams.proxy_access_type
        authentication          = $SystemParams.authentication
        use_svc_account_creds   = $SystemParams.use_svc_account_creds
        username                = $SystemParams.username
        password                = $SystemParams.password
    }

    $connection_string = ConvertTo-Json $connection_params -Compress -Depth 32

    if ($Global:MsExchangePSSession -and $connection_string -ne $Global:MsExchangeConnectionString) {
        Log info "MsExchangePSSession connection parameters changed"
        Close-MsExchangeSession
    }

    if ($Global:MsExchangePSSession -and $Global:MsExchangePSSession.State -ne 'Opened') {
        Log warn "MsExchangePSSession State is '$($Global:MsExchangePSSession.State)'"
        Close-MsExchangeSession
    }

    if ($Global:MsExchangePSSession) {
        #Log debug "Reusing MsExchangePSSession"
    }
    else {
        Log info "Opening MsExchangePSSession '$connection_string'"

        try {
            $protocol = if ($connection_params.use_secure_connection) { 'https' } else { 'http' }

            $new_ps_session_params = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = $protocol + '://' + $connection_params.server + '/PowerShell/'
                AllowRedirection  = $true
                Authentication    = $connection_params.authentication
            }

            if (-not $connection_params.use_svc_account_creds) {
                $new_ps_session_params.Credential = New-Object System.Management.Automation.PSCredential($connection_params.username, (ConvertTo-SecureString $connection_params.password -AsPlainText -Force))
            }

            if ($connection_params.skip_certificate_checks.Count -gt 0 -or $connection_params.use_proxy_server) {
                $new_ps_session_option_params = @{}

                if ($connection_params.skip_certificate_checks.Count -gt 0) {
                    $connection_params.skip_certificate_checks | ForEach-Object { $new_ps_session_option_params[$_] = $true }
                }

                if ($connection_params.use_proxy_server) {
                    $new_ps_session_option_params.ProxyAccessType = $connection_params.proxy_access_type
                }

                $new_ps_session_params.SessionOption = New-PSSessionOption @new_ps_session_option_params
            }

            $ps_session = New-PSSession @new_ps_session_params -WarningAction SilentlyContinue

            try {
                $ps_mod_info = Import-PSSession -AllowClobber -DisableNameChecking -Prefix 'MsExchange' -Session $ps_session
            }
            catch {
                Remove-PSSession -Session $ps_session -ErrorAction SilentlyContinue
                throw
            }

            $Global:MsExchangePSSession = $ps_session
            $Global:MsExchangeConnectionString = $connection_string
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }

        Log info "Done"
    }
}


function Close-MsExchangeSession {
    if ($Global:MsExchangePSSession) {
        Log info "Closing MsExchangePSSession"

        try {
            Remove-PSSession -Session $Global:MsExchangePSSession -ErrorAction SilentlyContinue
            $Global:MsExchangePSSession = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log info "Done"
    }
}


function Get-ClassMetaData {
    param (
        [string] $SystemParams,
        [string] $Class
    )

    @(
        @{
            name = 'properties'
            type = 'grid'
            label = 'Properties'
            description = 'Selected properties'
            table = @{
                rows = @( $Global:Properties.$Class | ForEach-Object {
                    @{
                        name = $_.name
                        usage_hint = @( @(
                            foreach ($key in $_.Keys) {
                                if ($key -notin @('default', 'idm', 'key')) { continue }

                                if ($key -eq 'idm') {
                                    $key.Toupper()
                                }
                                else {
                                    $key.Substring(0,1).Toupper() + $key.Substring(1)
                                }
                            }
                        ) | Sort-Object) -join ' | '
                    }
                })
                settings_grid = @{
                    selection = 'multiple'
                    key_column = 'name'
                    checkbox = $true
                    filter = $true
                    columns = @(
                        @{
                            name = 'name'
                            display_name = 'Name'
                        }
                        @{
                            name = 'usage_hint'
                            display_name = 'Usage hint'
                        }
                    )
                }
            }
            value = ($Global:Properties.$Class | Where-Object { $_.default }).name
        }
    )
}
