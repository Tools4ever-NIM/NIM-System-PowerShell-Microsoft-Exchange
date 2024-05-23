# version: 2.1
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
                tooltip = 'Name of Exchange server'
                value = 'ps.outlook.com'
            }
            @{
                name = 'use_secure_connection'
                type = 'checkbox'
                label = 'Use secure connection'
                tooltip = 'Using http or https'
                value = $false
            }
            @{
                name = 'skip_certificate_checks'
                type = 'checkgroup'
                label = 'Skip certificate checks'
                label_indent = $true
                tooltip = 'Skip the following certificate checks'
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
                tooltip = 'Behind a proxy server?'
                value = $false
            }
            @{
                name = 'proxy_access_type'
                type = 'combo'
                label = 'Proxy access type'
                label_indent = $true
                tooltip = 'Proxy access type'
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
                tooltip = 'Authentication method'
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
                tooltip = 'User account name to access exchange'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                tooltip = 'User account password to access exchange'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                tooltip = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                tooltip = ''
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
        @{ name = 'ActiveSyncAllowedDeviceIDs';                          options = @('set')                      }
        @{ name = 'ActiveSyncBlockedDeviceIDs';                          options = @('set')                      }
        @{ name = 'ActiveSyncDebugLogging';                              options = @('set')                      }
        @{ name = 'ActiveSyncEnabled';                                   options = @('default', 'set')           }
        @{ name = 'ActiveSyncMailboxPolicy';                             options = @('set')                      }
        @{ name = 'ActiveSyncMailboxPolicyIsDefaulted';                                                          }
        @{ name = 'ActiveSyncSuppressReadReceipt';                       options = @('set')                      }
        @{ name = 'DisplayName';                                         options = @('set')                      }
        @{ name = 'DistinguishedName';                                                                           }
        @{ name = 'ECPEnabled';                                          options = @('default', 'set')           }
        @{ name = 'EmailAddresses';                                      options = @('set')                      }
        @{ name = 'EwsAllowEntourage';                                   options = @('set')                      }
        @{ name = 'EwsAllowList';                                        options = @('set')                      }
        @{ name = 'EwsAllowMacOutlook';                                  options = @('set')                      }
        @{ name = 'EwsAllowOutlook';                                     options = @('set')                      }
        @{ name = 'EwsApplicationAccessPolicy';                          options = @('set')                      }
        @{ name = 'EwsBlockList';                                        options = @('set')                      }
        @{ name = 'EwsEnabled';                                          options = @('set')                      }
        @{ name = 'ExchangeVersion';                                                                             }
        @{ name = 'ExternalImapSettings';                                                                        }
        @{ name = 'ExternalPopSettings';                                                                         }
        @{ name = 'ExternalSmtpSettings';                                                                        }
        @{ name = 'Guid';                                                options = @('default', 'key')           }
        @{ name = 'HasActiveSyncDevicePartnership';                      options = @('set')                      }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                            options = @('default')                  }
        @{ name = 'ImapEnabled';                                         options = @('default', 'set')           }
        @{ name = 'ImapEnableExactRFC822Size';                           options = @('set')                      }
        @{ name = 'ImapForceICalForCalendarRetrievalOption';             options = @('set')                      }
        @{ name = 'ImapMessagesRetrievalMimeFormat';                     options = @('set')                      }
        @{ name = 'ImapSuppressReadReceipt';                             options = @('set')                      }
        @{ name = 'ImapUseProtocolDefaults';                             options = @('set')                      }
        @{ name = 'InternalImapSettings';                                                                        }
        @{ name = 'InternalPopSettings';                                                                         }
        @{ name = 'InternalSmtpSettings';                                                                        }
        @{ name = 'IsOptimizedForAccessibility';                         options = @('set')                      }
        @{ name = 'IsValid';                                             options = @('default')                  }
        @{ name = 'LegacyExchangeDN';                                                                            }
        @{ name = 'LinkedMasterAccount';                                 options = @('default')                  }
        @{ name = 'MAPIBlockOutlookExternalConnectivity';                options = @('set')                      }
        @{ name = 'MAPIBlockOutlookNonCachedMode';                       options = @('set')                      }
        @{ name = 'MAPIBlockOutlookRpcHttp';                             options = @('set')                      }
        @{ name = 'MAPIBlockOutlookVersions';                            options = @('set')                      }
        @{ name = 'MAPIEnabled';                                         options = @('set')                      }
        @{ name = 'MapiHttpEnabled';                                     options = @('set')                      }
        @{ name = 'Name';                                                options = @('set')                      }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'OWAEnabled';                                          options = @('default', 'set')           }
        @{ name = 'OWAforDevicesEnabled';                                options = @('set')                      }
        @{ name = 'OwaMailboxPolicy';                                    options = @('set')                      }
        @{ name = 'PopEnabled';                                          options = @('default', 'set')           }
        @{ name = 'PopEnableExactRFC822Size';                            options = @('set')                      }
        @{ name = 'PopForceICalForCalendarRetrievalOption';              options = @('set')                      }
        @{ name = 'PopMessageDeleteEnabled';                                                                     }
        @{ name = 'PopMessagesRetrievalMimeFormat';                      options = @('set')                      }
        @{ name = 'PopSuppressReadReceipt';                              options = @('set')                      }
        @{ name = 'PopUseProtocolDefaults';                              options = @('set')                      }
        @{ name = 'PrimarySmtpAddress';                                  options = @('default', 'set')           }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'PublicFolderClientAccess';                            options = @('set')                      }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'SamAccountName';                                      options = @('set')                      }
        @{ name = 'ServerLegacyDN';                                                                              }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'ShowGalAsDefaultView';                                options = @('set')                      }
        @{ name = 'UniversalOutlookEnabled';                             options = @('set')                      }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
    )

    Mailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                              options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                     options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';              options = @('set')                      }
        @{ name = 'AccountDisabled';                                     options = @('set')                      }
        @{ name = 'AddressBookPolicy';                                   options = @('enable', 'set')            }
        @{ name = 'AddressListMembership';                                                                       }
        @{ name = 'AdminDisplayVersion';                                                                         }
        @{ name = 'AdministrativeUnits';                                                                         }
        @{ name = 'AggregatedMailboxGuids';                                                                      }
        @{ name = 'Alias';                                               options = @('default', 'enable', 'set') }
        @{ name = 'AntispamBypassEnabled';                               options = @('set')                      }
        @{ name = 'ArbitrationMailbox';                                  options = @('set')                      }
        @{ name = 'ArchiveDatabase';                                     options = @('enable', 'set')            }
        @{ name = 'ArchiveDomain';                                       options = @('enable', 'set')            }
        @{ name = 'ArchiveGuid';                                         options = @('enable')                   }
        @{ name = 'ArchiveName';                                         options = @('default', 'enable', 'set') }
        @{ name = 'ArchiveQuota';                                        options = @('set')                      }
        @{ name = 'ArchiveRelease';                                                                              }
        @{ name = 'ArchiveState';                                        options = @('set')                      }
        @{ name = 'ArchiveStatus';                                                                               }
        @{ name = 'ArchiveWarningQuota';                                 options = @('set')                      }
        @{ name = 'AuditAdmin';                                          options = @('set')                      }
        @{ name = 'AuditDelegate';                                       options = @('set')                      }
        @{ name = 'AuditEnabled';                                        options = @('set')                      }
        @{ name = 'AuditLogAgeLimit';                                    options = @('set')                      }
        @{ name = 'AuditOwner';                                          options = @('set')                      }
        @{ name = 'AutoExpandingArchiveEnabled';                                                                 }
        @{ name = 'BypassModerationFromSendersOrMembers';                options = @('set')                      }
        @{ name = 'CalendarLoggingQuota';                                options = @('set')                      }
        @{ name = 'CalendarRepairDisabled';                              options = @('set')                      }
        @{ name = 'CalendarVersionStoreDisabled';                        options = @('set')                      }
        @{ name = 'ComplianceTagHoldApplied';                                                                    }
        @{ name = 'CustomAttribute1';                                    options = @('set')                      }
        @{ name = 'CustomAttribute2';                                    options = @('set')                      }
        @{ name = 'CustomAttribute3';                                    options = @('set')                      }
        @{ name = 'CustomAttribute4';                                    options = @('set')                      }
        @{ name = 'CustomAttribute5';                                    options = @('set')                      }
        @{ name = 'CustomAttribute6';                                    options = @('set')                      }
        @{ name = 'CustomAttribute7';                                    options = @('set')                      }
        @{ name = 'CustomAttribute8';                                    options = @('set')                      }
        @{ name = 'CustomAttribute9';                                    options = @('set')                      }
        @{ name = 'CustomAttribute10';                                   options = @('set')                      }
        @{ name = 'CustomAttribute11';                                   options = @('set')                      }
        @{ name = 'CustomAttribute12';                                   options = @('set')                      }
        @{ name = 'CustomAttribute13';                                   options = @('set')                      }
        @{ name = 'CustomAttribute14';                                   options = @('set')                      }
        @{ name = 'CustomAttribute15';                                   options = @('set')                      }
        @{ name = 'Database';                                            options = @('default', 'enable', 'set') }
        @{ name = 'DataEncryptionPolicy';                                options = @('set')                      }
        @{ name = 'DefaultPublicFolderMailbox';                          options = @('set')                      }
        @{ name = 'DelayHoldApplied';                                                                            }
        @{ name = 'DeliverToMailboxAndForward';                          options = @('set')                      }
        @{ name = 'DisabledArchiveDatabase';                                                                     }
        @{ name = 'DisabledArchiveGuid';                                                                         }
        @{ name = 'DisabledMailboxLocations';                                                                    }
        @{ name = 'DisableThrottling';                                   options = @('set')                      }
        @{ name = 'DisplayName';                                         options = @('default', 'enable', 'set') }
        @{ name = 'DistinguishedName';                                                                           }
        @{ name = 'DowngradeHighPriorityMessagesEnabled';                options = @('set')                      }
        @{ name = 'EffectivePublicFolderMailbox';                                                                }
        @{ name = 'ElcProcessingDisabled';                               options = @('set')                      }
        @{ name = 'EmailAddresses';                                      options = @('default', 'set')           }
        @{ name = 'EmailAddressPolicyEnabled';                           options = @('set')                      }
        @{ name = 'EndDateForRetentionHold';                             options = @('set')                      }
        @{ name = 'ExchangeGuid';                                                                                }
        @{ name = 'ExchangeUserAccountControl';                                                                  }
        @{ name = 'ExchangeSecurityDescriptor';                                                                  }
        @{ name = 'ExchangeVersion';                                                                             }
        @{ name = 'ExtensionCustomAttribute1';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                           options = @('set')                      }
        @{ name = 'Extensions';                                                                                  }
        @{ name = 'ExternalDirectoryObjectId';                                                                   }
        @{ name = 'ExternalOofOptions';                                  options = @('set')                      }
        @{ name = 'ForwardingAddress';                                   options = @('set')                      }
        @{ name = 'ForwardingSmtpAddress';                               options = @('set')                      }
        @{ name = 'GeneratedOfflineAddressBooks';                                                                }
        @{ name = 'GrantSendOnBehalfTo';                                 options = @('set')                      }
        @{ name = 'Guid';                                                options = @('default', 'key')           }
        @{ name = 'HasPicture';                                                                                  }
        @{ name = 'HasSnackyAppData';                                                                            }
        @{ name = 'HasSpokenName';                                                                               }
        @{ name = 'HiddenFromAddressListsEnabled';                       options = @('set')                      }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                                                                    }
        @{ name = 'ImListMigrationCompleted';                            options = @('set')                      }
        @{ name = 'ImmutableId';                                         options = @('set')                      }
        @{ name = 'InactiveMailboxRetireTime';                                                                   }
        @{ name = 'IncludeInGarbageCollection';                                                                  }
        @{ name = 'InPlaceHolds';                                                                                }
        @{ name = 'IsDirSynced';                                                                                 }
        @{ name = 'IsExcludedFromServingHierarchy';                      options = @('set')                      }
        @{ name = 'IsHierarchyReady';                                    options = @('set')                      }
        @{ name = 'IsHierarchySyncEnabled';                              options = @('set')                      }
        @{ name = 'IsInactiveMailbox';                                                                           }
        @{ name = 'IsLinked';                                                                                    }
        @{ name = 'IsMachineToPersonTextMessagingEnabled';                                                       }
        @{ name = 'IsMailboxEnabled';                                                                            }
        @{ name = 'IsPersonToPersonTextMessagingEnabled';                                                        }
        @{ name = 'IsResource';                                                                                  }
        @{ name = 'IsRootPublicFolderMailbox';                                                                   }
        @{ name = 'IsShared';                                                                                    }
        @{ name = 'IsSoftDeletedByDisable';                                                                      }
        @{ name = 'IsSoftDeletedByRemove';                                                                       }
        @{ name = 'IssueWarningQuota';                                   options = @('set')                      }
        @{ name = 'IsValid';                                                                                     }
        @{ name = 'JournalArchiveAddress';                               options = @('set')                      }
        @{ name = 'Languages';                                           options = @('set')                      }
        @{ name = 'LastExchangeChangedTime';                                                                     }
        @{ name = 'LegacyExchangeDN';                                                                            }
        @{ name = 'LinkedMasterAccount';                                 options = @('enable', 'set')            }
        @{ name = 'LitigationHoldDate';                                  options = @('set')                      }
        @{ name = 'LitigationHoldDuration';                              options = @('set')                      }
        @{ name = 'LitigationHoldEnabled';                               options = @('set')                      }
        @{ name = 'LitigationHoldOwner';                                 options = @('set')                      }
        @{ name = 'MailboxContainerGuid';                                                                        }
        @{ name = 'MailboxLocations';                                                                            }
        @{ name = 'MailboxMoveBatchName';                                                                        }
        @{ name = 'MailboxMoveFlags';                                                                            }
        @{ name = 'MailboxMoveRemoteHostName';                                                                   }
        @{ name = 'MailboxMoveSourceMDB';                                                                        }
        @{ name = 'MailboxMoveStatus';                                                                           }
        @{ name = 'MailboxMoveTargetMDB';                                                                        }
        @{ name = 'MailboxPlan';                                                                                 }
        @{ name = 'MailboxProvisioningConstraint';                                                               }
        @{ name = 'MailboxProvisioningPreferences';                                                              }
        @{ name = 'MailboxRegion';                                       options = @('set')                      }
        @{ name = 'MailboxRegionLastUpdateTime';                                                                 }
        @{ name = 'MailboxRelease';                                                                              }
        @{ name = 'MailTip';                                             options = @('set')                      }
        @{ name = 'MailTipTranslations';                                 options = @('set')                      }
        @{ name = 'ManagedFolderMailboxPolicy';                          options = @('enable', 'set')            }
        @{ name = 'MaxBlockedSenders';                                   options = @('set')                      }
        @{ name = 'MaxReceiveSize';                                      options = @('set')                      }
        @{ name = 'MaxSafeSenders';                                      options = @('set')                      }
        @{ name = 'MaxSendSize';                                         options = @('set')                      }
        @{ name = 'MessageCopyForSendOnBehalfEnabled';                   options = @('set')                      }
        @{ name = 'MessageCopyForSentAsEnabled';                         options = @('set')                      }
        @{ name = 'MessageTrackingReadStatusEnabled';                                                            }
        @{ name = 'MicrosoftOnlineServicesID';                           options = @('set')                      }
        @{ name = 'ModeratedBy';                                         options = @('set')                      }
        @{ name = 'ModerationEnabled';                                   options = @('set')                      }
        @{ name = 'Name';                                                options = @('set')                      }
        @{ name = 'NetID';                                                                                       }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'Office';                                              options = @('set')                      }
        @{ name = 'OfflineAddressBook';                                  options = @('set')                      }
        @{ name = 'OrganizationalUnit';                                                                          }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'OrphanSoftDeleteTrackingTime';                                                                }
        @{ name = 'PersistedCapabilities';                                                                       }
        @{ name = 'PoliciesExcluded';                                                                            }
        @{ name = 'PoliciesIncluded';                                                                            }
        @{ name = 'PrimarySmtpAddress';                                  options = @('default', 'enable', 'set') }
        @{ name = 'ProhibitSendQuota';                                   options = @('set')                      }
        @{ name = 'ProhibitSendReceiveQuota';                            options = @('set')                      }
        @{ name = 'ProtocolSettings';                                                                            }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'QueryBaseDN';                                         options = @('set')                      }
        @{ name = 'QueryBaseDNRestrictionEnabled';                                                               }
        @{ name = 'RecipientLimits';                                     options = @('set')                      }
        @{ name = 'RecipientType';                                                                               }
        @{ name = 'RecipientTypeDetails';                                                                        }
        @{ name = 'ReconciliationId';                                                                            }
        @{ name = 'RecoverableItemsQuota';                               options = @('set')                      }
        @{ name = 'RecoverableItemsWarningQuota';                        options = @('set')                      }
        @{ name = 'RejectMessagesFrom';                                  options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                         options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';                  options = @('set')                      }
        @{ name = 'RemoteAccountPolicy';                                                                         }
        @{ name = 'RemoteRecipientType';                                 options = @('set')                      }
        @{ name = 'RequireSenderAuthenticationEnabled';                  options = @('set')                      }
        @{ name = 'ResetPasswordOnNextLogon';                            options = @('set')                      }
        @{ name = 'ResourceCapacity';                                    options = @('set')                      }
        @{ name = 'ResourceCustom';                                      options = @('set')                      }
        @{ name = 'ResourceType';                                                                                }
        @{ name = 'RetainDeletedItemsFor';                               options = @('set')                      }
        @{ name = 'RetainDeletedItemsUntilBackup';                       options = @('set')                      }
        @{ name = 'RetentionComment';                                    options = @('set')                      }
        @{ name = 'RetentionHoldEnabled';                                options = @('set')                      }
        @{ name = 'RetentionPolicy';                                     options = @('enable', 'set')            }
        @{ name = 'RetentionUrl';                                        options = @('set')                      }
        @{ name = 'RoleAssignmentPolicy';                                options = @('enable', 'set')            }
        @{ name = 'RoomMailboxAccountEnabled';                                                                   }
        @{ name = 'RulesQuota';                                          options = @('set')                      }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'SamAccountName';                                      options = @('set')                      }
        @{ name = 'SCLDeleteEnabled';                                    options = @('set')                      }
        @{ name = 'SCLDeleteThreshold';                                  options = @('set')                      }
        @{ name = 'SCLJunkEnabled';                                      options = @('set')                      }
        @{ name = 'SCLJunkThreshold';                                    options = @('set')                      }
        @{ name = 'SCLQuarantineEnabled';                                options = @('set')                      }
        @{ name = 'SCLQuarantineThreshold';                              options = @('set')                      }
        @{ name = 'SCLRejectEnabled';                                    options = @('set')                      }
        @{ name = 'SCLRejectThreshold';                                  options = @('set')                      }
        @{ name = 'SendModerationNotifications';                         options = @('set')                      }
        @{ name = 'ServerLegacyDN';                                                                              }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'SharingPolicy';                                       options = @('set')                      }
        @{ name = 'SiloName';                                                                                    }
        @{ name = 'SimpleDisplayName';                                   options = @('set')                      }
        @{ name = 'SingleItemRecoveryEnabled';                           options = @('set')                      }
        @{ name = 'SKUAssigned';                                                                                 }
        @{ name = 'SourceAnchor';                                                                                }
        @{ name = 'StartDateForRetentionHold';                           options = @('set')                      }
        @{ name = 'StsRefreshTokensValidFrom';                           options = @('set')                      }
        @{ name = 'ThrottlingPolicy';                                    options = @('set')                      }
        @{ name = 'UMDtmfMap';                                           options = @('set')                      }
        @{ name = 'UMEnabled';                                                                                   }
        @{ name = 'UnifiedMailbox';                                                                              }
        @{ name = 'UsageLocation';                                                                               }
        @{ name = 'UseDatabaseQuotaDefaults';                            options = @('set')                      }
        @{ name = 'UseDatabaseRetentionDefaults';                        options = @('set')                      }
        @{ name = 'UserCertificate';                                     options = @('set')                      }
        @{ name = 'UserPrincipalName';                                   options = @('set')                      }
        @{ name = 'UserSMimeCertificate';                                options = @('set')                      }
        @{ name = 'WasInactiveMailbox';                                                                          }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
        @{ name = 'WhenMailboxCreated';                                                                          }
        @{ name = 'WhenSoftDeleted';                                                                             }
        @{ name = 'WindowsEmailAddress';                                 options = @('set')                      }
        @{ name = 'WindowsLiveID';                                                                               }
    )

    MailboxDatabase = @(
        @{ name = 'ActivationPreference';                                                                        }
        @{ name = 'AdminDisplayName';                                                                            }
        @{ name = 'AdminDisplayVersion';                                                                         }
        @{ name = 'AdministrativeGroup';                                                                         }
        @{ name = 'AllDatabaseCopies';                                                                           }
        @{ name = 'AllowFileRestore';                                                                            }
        @{ name = 'AutoDagExcludeFromMonitoring';                                                                }
        @{ name = 'AutoDatabaseMountDial';                                                                       }
        @{ name = 'AvailableNewMailboxSpace';                                                                    }
        @{ name = 'BackgroundDatabaseMaintenance';                                                               }
        @{ name = 'BackgroundDatabaseMaintenanceDelay';                                                          }
        @{ name = 'BackgroundDatabaseMaintenanceSerialization';                                                  }
        @{ name = 'BackupInProgress';                                                                            }
        @{ name = 'CachedClosedTables';                                                                          }
        @{ name = 'CachePriority';                                                                               }
        @{ name = 'CafeEndpoints';                                                                               }
        @{ name = 'CalendarLoggingQuota';                                                                        }
        @{ name = 'CircularLoggingEnabled';                                                                      }
        @{ name = 'CreationSchemaVersion';                                                                       }
        @{ name = 'CurrentSchemaVersion';                                                                        }
        @{ name = 'DatabaseCopies';                                                                              }
        @{ name = 'DatabaseCreated';                                                                             }
        @{ name = 'DatabaseExtensionSize';                                                                       }
        @{ name = 'DatabaseGroup';                                                                               }
        @{ name = 'DatabaseSize';                                        options = @('default')                  }
        @{ name = 'DataMoveReplicationConstraint';                                                               }
        @{ name = 'DeletedItemRetention';                                                                        }
        @{ name = 'Description';                                         options = @('default')                  }
        @{ name = 'DistinguishedName';                                                                           }
        @{ name = 'DumpsterServersNotAvailable';                                                                 }
        @{ name = 'DumpsterStatistics';                                                                          }
        @{ name = 'EdbFilePath';                                                                                 }
        @{ name = 'EventHistoryRetentionPeriod';                                                                 }
        @{ name = 'ExchangeLegacyDN';                                                                            }
        @{ name = 'ExchangeVersion';                                     options = @('default')                  }
        @{ name = 'Guid';                                                                                        }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                            options = @('default', 'key')           }
        @{ name = 'IndexEnabled';                                                                                }
        @{ name = 'InvalidDatabaseCopies';                                                                       }
        @{ name = 'IsExcludedFromInitialProvisioning';                                                           }
        @{ name = 'IsExcludedFromProvisioning';                                                                  }
        @{ name = 'IsExcludedFromProvisioningBy';                                                                }
        @{ name = 'IsExcludedFromProvisioningByOperator';                                                        }
        @{ name = 'IsExcludedFromProvisioningBySchemaVersionMonitoring';                                         }
        @{ name = 'IsExcludedFromProvisioningBySpaceMonitoring';                                                 }
        @{ name = 'IsExcludedFromProvisioningDueToLogicalCorruption';                                            }
        @{ name = 'IsExcludedFromProvisioningForDraining';                                                       }
        @{ name = 'IsExcludedFromProvisioningReason';                                                            }
        @{ name = 'IsMailboxDatabase';                                   options = @('default')                  }
        @{ name = 'IsPublicFolderDatabase';                              options = @('default')                  }
        @{ name = 'IssueWarningQuota';                                   options = @('default')                  }
        @{ name = 'IsSuspendedFromProvisioning';                         options = @('default')                  }
        @{ name = 'IsValid';                                             options = @('default')                  }
        @{ name = 'JournalRecipient';                                                                            }
        @{ name = 'LastCopyBackup';                                                                              }
        @{ name = 'LastDifferentialBackup';                                                                      }
        @{ name = 'LastFullBackup';                                                                              }
        @{ name = 'LastIncrementalBackup';                                                                       }
        @{ name = 'LogBuffers';                                                                                  }
        @{ name = 'LogCheckpointDepth';                                                                          }
        @{ name = 'LogFilePrefix';                                                                               }
        @{ name = 'LogFileSize';                                                                                 }
        @{ name = 'LogFolderPath';                                                                               }
        @{ name = 'MailboxLoadBalanceEnabled';                                                                   }
        @{ name = 'MailboxLoadBalanceMaximumEdbFileSize';                                                        }
        @{ name = 'MailboxLoadBalanceOverloadedThreshold';                                                       }
        @{ name = 'MailboxLoadBalanceRelativeLoadCapacity';                                                      }
        @{ name = 'MailboxLoadBalanceUnderloadedThreshold';                                                      }
        @{ name = 'MailboxProvisioningAttributes';                                                               }
        @{ name = 'MailboxRetention';                                                                            }
        @{ name = 'MaintenanceSchedule';                                                                         }
        @{ name = 'MasterServerOrAvailabilityGroup';                                                             }
        @{ name = 'MasterType';                                                                                  }
        @{ name = 'MaximumBackgroundDatabaseMaintenanceInterval';                                                }
        @{ name = 'MaximumCursors';                                                                              }
        @{ name = 'MaximumOpenTables';                                                                           }
        @{ name = 'MaximumPreReadPages';                                                                         }
        @{ name = 'MaximumReplayPreReadPages';                                                                   }
        @{ name = 'MaximumSessions';                                                                             }
        @{ name = 'MaximumTemporaryTables';                                                                      }
        @{ name = 'MaximumVersionStorePages';                                                                    }
        @{ name = 'MCDBAvailableSpace';                                                                          }
        @{ name = 'MCDBSize';                                                                                    }
        @{ name = 'MetaCacheDatabaseFilePath';                                                                   }
        @{ name = 'MetaCacheDatabaseFolderPath';                                                                 }
        @{ name = 'MetaCacheDatabaseMaxCapacityInBytes';                                                         }
        @{ name = 'MetaCacheDatabaseMountpointFolderPath';                                                       }
        @{ name = 'MetaCacheDatabaseRootFolderPath';                                                             }
        @{ name = 'MetaCacheDatabaseVolumesRootFolderPath';                                                      }
        @{ name = 'MimimumBackgroundDatabaseMaintenanceInterval';                                                }
        @{ name = 'MountAtStartup';                                                                              }
        @{ name = 'Mounted';                                                                                     }
        @{ name = 'MountedOnServer';                                                                             }
        @{ name = 'Name';                                                                                        }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'OfflineAddressBook';                                                                          }
        @{ name = 'Organization';                                                                                }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginalDatabase';                                                                            }
        @{ name = 'OriginatingServer';                                   options = @('default')                  }
        @{ name = 'PreferredVersionStorePages';                                                                  }
        @{ name = 'ProhibitSendQuota';                                                                           }
        @{ name = 'ProhibitSendReceiveQuota';                                                                    }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'PublicFolderDatabase';                                                                        }
        @{ name = 'QuotaNotificationSchedule';                                                                   }
        @{ name = 'RecoverableItemsQuota';                                                                       }
        @{ name = 'RecoverableItemsWarningQuota';                                                                }
        @{ name = 'Recovery';                                                                                    }
        @{ name = 'ReplayBackgroundDatabaseMaintenance';                                                         }
        @{ name = 'ReplayBackgroundDatabaseMaintenanceDelay';                                                    }
        @{ name = 'ReplayCachePriority';                                                                         }
        @{ name = 'ReplayCheckpointDepth';                                                                       }
        @{ name = 'ReplayLagTimes';                                                                              }
        @{ name = 'ReplicationType';                                                                             }
        @{ name = 'RequestedSchemaVersion';                                                                      }
        @{ name = 'RetainDeletedItemsUntilBackup';                                                               }
        @{ name = 'RpcClientAccessServer';                                                                       }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'Server';                                                                                      }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'Servers';                                                                                     }
        @{ name = 'SnapshotLastCopyBackup';                                                                      }
        @{ name = 'SnapshotLastDifferentialBackup';                                                              }
        @{ name = 'SnapshotLastFullBackup';                                                                      }
        @{ name = 'SnapshotLastIncrementalBackup';                                                               }
        @{ name = 'TemporaryDataFolderPath';                                                                     }
        @{ name = 'TruncationLagTimes';                                                                          }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
        @{ name = 'WorkerProcessId';                                                                             }
    )

    MailboxPermission = @(
        @{ name = 'ExchangeGuid';                                        options = @('default','key')            }         
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'AccessRights';                                        options = @('default')                  }
        @{ name = 'Deny';                                                options = @('default')                  }
        @{ name = 'InheirtanceType';                                     options = @('default')                  }
        @{ name = 'User';                                                options = @('default')                  }
        @{ name = 'Identity';                                            options = @('default')                  }
        @{ name = 'IsInherited';                                         options = @('default')                  }
        @{ name = 'True';                                                options = @('default')                  }
        @{ name = 'Unchanged';                                           options = @('default')                  }
    )

    MailboxStatistics = @(
        @{ name = 'AssociatedItemCount';                                                                         }
        @{ name = 'AttachmentTableAvailableSize';                                                                }
        @{ name = 'AttachmentTableTotalSize';                                                                    }
        @{ name = 'CurrentSchemaVersion';                                                                        }
        @{ name = 'Database';                                                                                    }
        @{ name = 'DatabaseName';                                        options = @('default')                  }
        @{ name = 'DatabaseIssueWarningQuota';                                                                   }
        @{ name = 'DatabaseProhibitSendQuota';                                                                   }
        @{ name = 'DatabaseProhibitSendReceiveQuota';                                                            }
        @{ name = 'DeletedItemCount';                                                                            }
        @{ name = 'DisconnectDate';                                      options = @('default')                  }
        @{ name = 'DisconnectReason';                                    options = @('default')                  }
        @{ name = 'DisplayName';                                         options = @('default')                  }
        @{ name = 'DumpsterMessagesPerFolderCountReceiveQuota';                                                  }
        @{ name = 'DumpsterMessagesPerFolderCountWarningQuota';                                                  }
        @{ name = 'ExternalDirectoryOrganizationId';                                                             }
        @{ name = 'FolderHierarchyChildrenCountReceiveQuota';                                                    }
        @{ name = 'FolderHierarchyChildrenCountWarningQuota';                                                    }
        @{ name = 'FolderHierarchyDepthReceiveQuota';                                                            }
        @{ name = 'FolderHierarchyDepthWarningQuota';                                                            }
        @{ name = 'FoldersCountReceiveQuota';                                                                    }
        @{ name = 'FoldersCountWarningQuota';                                                                    }
        @{ name = 'Identity';                                            options = @('default', 'key')           }
        @{ name = 'IsArchiveMailbox';                                    options = @('default')                  }
        @{ name = 'IsDatabaseCopyActive';                                                                        }
        @{ name = 'IsMoveDestination';                                                                           }
        @{ name = 'IsQuarantined';                                       options = @('default')                  }
        @{ name = 'IsValid';                                                                                     }
        @{ name = 'ItemCount';                                                                                   }
        @{ name = 'LastLoggedOnUserAccount';                             options = @('default')                  }
        @{ name = 'LastLogoffTime';                                                                              }
        @{ name = 'LastLogonTime';                                       options = @('default')                  }
        @{ name = 'LegacyDN';                                            options = @('default')                  }
        @{ name = 'MailboxGuid';                                         options = @('default')                  }
        @{ name = 'MailboxMessagesPerFolderCountReceiveQuota';                                                   }
        @{ name = 'MailboxMessagesPerFolderCountWarningQuota';                                                   }
        @{ name = 'MailboxTableIdentifier';                                                                      }
        @{ name = 'MailboxType';                                                                                 }
        @{ name = 'MapiIdentity';                                                                                }
        @{ name = 'MessageTableAvailableSize';                                                                   }
        @{ name = 'MessageTableTotalSize';                                                                       }
        @{ name = 'MoveHistory';                                                                                 }
        @{ name = 'NamedPropertiesCountQuota';                                                                   }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'OtherTablesAvailableSize';                                                                    }
        @{ name = 'OtherTablesTotalSize';                                                                        }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'QuarantineDescription';                                                                       }
        @{ name = 'QuarantineEnd';                                                                               }
        @{ name = 'QuarantineFileVersion';                                                                       }
        @{ name = 'QuarantineLastCrash';                                                                         }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'ServerName';                                                                                  }
        @{ name = 'StorageLimitStatus';                                                                          }
        @{ name = 'TotalDeletedItemSize';                                                                        }
        @{ name = 'TotalItemSize';                                       options = @('default')                  }
    )

    RemoteMailbox = @(
        @{ name = 'AcceptMessagesOnlyFrom';                              options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                     options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';              options = @('set')                      }
        @{ name = 'AccountDisabled';                                                                             }
        @{ name = 'AddressListMembership';                                                                       }
        @{ name = 'AggregatedMailboxGuids';                                                                      }
        @{ name = 'Alias';                                               options = @('default', 'enable', 'set') }
        @{ name = 'ArbitrationMailbox';                                                                          }
        @{ name = 'ArchiveDatabase';                                                                             }
        @{ name = 'ArchiveGuid';                                         options = @('set')                      }
        @{ name = 'ArchiveName';                                         options = @('default', 'enable', 'set') }
        @{ name = 'ArchiveQuota';                                                                                }
        @{ name = 'ArchiveRelease';                                                                              }
        @{ name = 'ArchiveState';                                                                                }
        @{ name = 'ArchiveStatus';                                                                               }
        @{ name = 'ArchiveWarningQuota';                                                                         }
        @{ name = 'BypassModerationFromSendersOrMembers';                options = @('set')                      }
        @{ name = 'CalendarVersionStoreDisabled';                                                                }
        @{ name = 'CustomAttribute1';                                    options = @('set')                      }
        @{ name = 'CustomAttribute2';                                    options = @('set')                      }
        @{ name = 'CustomAttribute3';                                    options = @('set')                      }
        @{ name = 'CustomAttribute4';                                    options = @('set')                      }
        @{ name = 'CustomAttribute5';                                    options = @('set')                      }
        @{ name = 'CustomAttribute6';                                    options = @('set')                      }
        @{ name = 'CustomAttribute7';                                    options = @('set')                      }
        @{ name = 'CustomAttribute8';                                    options = @('set')                      }
        @{ name = 'CustomAttribute9';                                    options = @('set')                      }
        @{ name = 'CustomAttribute10';                                   options = @('set')                      }
        @{ name = 'CustomAttribute11';                                   options = @('set')                      }
        @{ name = 'CustomAttribute12';                                   options = @('set')                      }
        @{ name = 'CustomAttribute13';                                   options = @('set')                      }
        @{ name = 'CustomAttribute14';                                   options = @('set')                      }
        @{ name = 'CustomAttribute15';                                   options = @('set')                      }
        @{ name = 'DeliverToMailboxAndForward';                                                                  }
        @{ name = 'DisabledArchiveDatabase';                                                                     }
        @{ name = 'DisabledArchiveGuid';                                                                         }
        @{ name = 'DisplayName';                                         options = @('enable', 'set')            }
        @{ name = 'DistinguishedName';                                   options = @('default')                  }
        @{ name = 'EmailAddresses';                                      options = @('set')                      }
        @{ name = 'EmailAddressPolicyEnabled';                           options = @('set')                      }
        @{ name = 'EndDateForRetentionHold';                                                                     }
        @{ name = 'ExchangeGuid';                                        options = @('set')                      }
        @{ name = 'ExchangeVersion';                                                                             }
        @{ name = 'ExchangeUserAccountControl';                                                                  }
        @{ name = 'ExtensionCustomAttribute1';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                           options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                           options = @('set')                      }
        @{ name = 'Extensions';                                                                                  }
        @{ name = 'ExternalDirectoryObjectId';                                                                   }
        @{ name = 'ForwardingAddress';                                                                           }
        @{ name = 'GrantSendOnBehalfTo';                                 options = @('set')                      }
        @{ name = 'Guid';                                                options = @('default', 'key')           }
        @{ name = 'HasPicture';                                                                                  }
        @{ name = 'HasSpokenName';                                                                               }
        @{ name = 'HiddenFromAddressListsEnabled';                       options = @('set')                      }
        @{ name = 'Id';                                                  options = @('default')                  }
        @{ name = 'Identity';                                                                                    }
        @{ name = 'ImmutableId';                                         options = @('set')                      }
        @{ name = 'InPlaceHolds';                                                                                }
        @{ name = 'IsSoftDeletedByDisable';                                                                      }
        @{ name = 'IsSoftDeletedByRemove';                                                                       }
        @{ name = 'IsValid';                                                                                     }
        @{ name = 'JournalArchiveAddress';                                                                       }
        @{ name = 'LastExchangeChangedTime';                                                                     }
        @{ name = 'LegacyExchangeDN';                                                                            }
        @{ name = 'LitigationHoldDate';                                                                          }
        @{ name = 'LitigationHoldEnabled';                                                                       }
        @{ name = 'LitigationHoldOwner';                                                                         }
        @{ name = 'MailboxContainerGuid';                                                                        }
        @{ name = 'MailboxLocations';                                                                            }
        @{ name = 'MailboxMoveBatchName';                                                                        }
        @{ name = 'MailboxMoveFlags';                                                                            }
        @{ name = 'MailboxMoveRemoteHostName';                                                                   }
        @{ name = 'MailboxMoveSourceMDB';                                                                        }
        @{ name = 'MailboxMoveStatus';                                                                           }
        @{ name = 'MailboxMoveTargetMDB';                                                                        }
        @{ name = 'MailboxProvisioningConstraint';                                                               }
        @{ name = 'MailboxProvisioningPreferences';                                                              }
        @{ name = 'MailboxRelease';                                                                              }
        @{ name = 'MailTip';                                             options = @('set')                      }
        @{ name = 'MailTipTranslations';                                 options = @('set')                      }
        @{ name = 'MaxReceiveSize';                                                                              }
        @{ name = 'MaxSendSize';                                                                                 }
        @{ name = 'ModeratedBy';                                         options = @('set')                      }
        @{ name = 'ModerationEnabled';                                   options = @('set')                      }
        @{ name = 'Name';                                                options = @('default', 'set')           }
        @{ name = 'ObjectCategory';                                                                              }
        @{ name = 'ObjectClass';                                                                                 }
        @{ name = 'ObjectState';                                                                                 }
        @{ name = 'OnPremisesOrganizationalUnit';                                                                }
        @{ name = 'OrganizationId';                                                                              }
        @{ name = 'OriginatingServer';                                                                           }
        @{ name = 'PersistedCapabilities';                                                                       }
        @{ name = 'PoliciesExcluded';                                                                            }
        @{ name = 'PoliciesIncluded';                                                                            }
        @{ name = 'ProtocolSettings';                                                                            }
        @{ name = 'PrimarySmtpAddress';                                  options = @('default', 'enable', 'set') }
        @{ name = 'PSComputerName';                                                                              }
        @{ name = 'PSShowComputerName';                                                                          }
        @{ name = 'RecipientLimits';                                                                             }
        @{ name = 'RecipientType';                                                                               }
        @{ name = 'RecipientTypeDetails';                                                                        }
        @{ name = 'RecoverableItemsQuota';                               options = @('set')                      }
        @{ name = 'RecoverableItemsWarningQuota';                        options = @('set')                      }
        @{ name = 'RejectMessagesFrom';                                  options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                         options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';                  options = @('set')                      }
        @{ name = 'RemoteRecipientType';                                                                         }
        @{ name = 'RemoteRoutingAddress';                                options = @('default', 'enable', 'set') }
        @{ name = 'RequireSenderAuthenticationEnabled';                  options = @('set')                      }
        @{ name = 'ResetPasswordOnNextLogon';                            options = @('set')                      }
        @{ name = 'RetainDeletedItemsFor';                                                                       }
        @{ name = 'RetentionComment';                                                                            }
        @{ name = 'RetentionHoldEnabled';                                                                        }
        @{ name = 'RetentionUrl';                                                                                }
        @{ name = 'RunspaceId';                                                                                  }
        @{ name = 'SamAccountName';                                      options = @('set')                      }
        @{ name = 'SendModerationNotifications';                         options = @('set')                      }
        @{ name = 'SimpleDisplayName';                                                                           }
        @{ name = 'SingleItemRecoveryEnabled';                                                                   }
        @{ name = 'StartDateForRetentionHold';                                                                   }
        @{ name = 'StsRefreshTokensValidFrom';                                                                   }
        @{ name = 'UMDtmfMap';                                                                                   }
        @{ name = 'UserCertificate';                                                                             }
        @{ name = 'UserPrincipalName';                                   options = @('set')                      }
        @{ name = 'UserSMimeCertificate';                                                                        }
        @{ name = 'WhenChanged';                                                                                 }
        @{ name = 'WhenChangedUTC';                                                                              }
        @{ name = 'WhenCreated';                                                                                 }
        @{ name = 'WhenCreatedUTC';                                                                              }
        @{ name = 'WhenMailboxCreated';                                                                          }
        @{ name = 'WhenSoftDeleted';                                                                             }
        @{ name = 'WindowsEmailAddress';                                 options = @('set')                      }
    )

    MailUser = @(
        @{ name = 'AcceptMessagesOnlyFrom';                          options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromDLMembers';                          options = @('set')                      }
        @{ name = 'AcceptMessagesOnlyFromSendersOrMembers';                          options = @('set')                      }
        @{ name = 'AccountDisabled';             }
        @{ name = 'AddressListMembership';        }
        @{ name = 'AdministrativeUnits';           }
        @{ name = 'AggregatedMailboxGuids';         options = @('set')                      }
        @{ name = 'Alias';                          options = @('default','set','enable')                      }
        @{ name = 'ArbitrationMailbox';                          options = @('set')                      }
        @{ name = 'ArchiveDatabase';                }
        @{ name = 'ArchiveGuid';                                               }
        @{ name = 'ArchiveName';                          options = @('default','set')                      }
        @{ name = 'ArchiveQuota';         }
        @{ name = 'ArchiveRelease';        }
        @{ name = 'ArchiveStatus';          }
        @{ name = 'ArchiveWarningQuota';     }
        @{ name = 'BypassModerationFromSendersOrMembers';                          options = @('set')                      }
        @{ name = 'CalendarVersionStoreDisabled';           }
        @{ name = 'ComplianceTagHoldApplied';                }
        @{ name = 'CustomAttribute1';                          options = @('set')                      }
        @{ name = 'CustomAttribute10';                          options = @('set')                      }
        @{ name = 'CustomAttribute11';                          options = @('set')                      }
        @{ name = 'CustomAttribute12';                          options = @('set')                      }
        @{ name = 'CustomAttribute13';                          options = @('set')                      }
        @{ name = 'CustomAttribute14';                          options = @('set')                      }
        @{ name = 'CustomAttribute15';                          options = @('set')                      }
        @{ name = 'CustomAttribute2';                          options = @('set')                      }
        @{ name = 'CustomAttribute3';                          options = @('set')                      }
        @{ name = 'CustomAttribute4';                          options = @('set')                      }
        @{ name = 'CustomAttribute5';                          options = @('set')                      }
        @{ name = 'CustomAttribute6';                          options = @('set')                      }
        @{ name = 'CustomAttribute7';                          options = @('set')                      }
        @{ name = 'CustomAttribute8';                          options = @('set')                      }
        @{ name = 'CustomAttribute9';                          options = @('set')                      }
        @{ name = 'DataEncryptionPolicy';               }
        @{ name = 'DelayHoldApplied';                    }
        @{ name = 'DeliverToMailboxAndForward';           }
        @{ name = 'DisabledArchiveDatabase';               }
        @{ name = 'DisabledArchiveGuid';                   }
        @{ name = 'DisplayName';                          options = @('default','set','enable')                      }
        @{ name = 'DistinguishedName';                }
        @{ name = 'EmailAddressPolicyEnabled';                          options = @('set')                      }
        @{ name = 'EmailAddresses';                          options = @('default','set')                      }
        @{ name = 'EndDateForRetentionHold';                          options = @('default','set')                      }
        @{ name = 'ExchangeGuid';                                               }
        @{ name = 'ExchangeUserAccountControl';         }
        @{ name = 'ExchangeVersion';                    }
        @{ name = 'ExtensionCustomAttribute1';                          options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute2';                          options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute3';                          options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute4';                          options = @('set')                      }
        @{ name = 'ExtensionCustomAttribute5';                          options = @('set')                      }
        @{ name = 'Extensions';                     }
        @{ name = 'ExternalDirectoryObjectId';       }
        @{ name = 'ExternalEmailAddress';                          options = @('default','set','enable')                      }
        @{ name = 'ForwardingAddress';                }
        @{ name = 'GrantSendOnBehalfTo';                          options = @('set')                      }
        @{ name = 'GuestInfo';                      }
        @{ name = 'Guid';                          options = @('key')                      }
        @{ name = 'HasPicture';                   }
        @{ name = 'HasSpokenName';                 }
        @{ name = 'HiddenFromAddressListsEnabled';                          options = @('set')                      }
        @{ name = 'Id';                         options = @('default')                      }
        @{ name = 'Identity';                          }
        @{ name = 'ImmutableId';                                               options = @('set')                      }
        @{ name = 'InPlaceHolds';                }
        @{ name = 'IsDirSynced';                  }
        @{ name = 'IsSoftDeletedByDisable';        }
        @{ name = 'IsSoftDeletedByRemove';          }
        @{ name = 'IsValid';                       }
        @{ name = 'IssueWarningQuota';              }
        @{ name = 'JournalArchiveAddress';           }
        @{ name = 'LastExchangeChangedTime';          }
        @{ name = 'LegacyExchangeDN';                  }
        @{ name = 'LitigationHoldDate';                 }
        @{ name = 'LitigationHoldEnabled';               }
        @{ name = 'LitigationHoldOwner';       }
        @{ name = 'MacAttachmentFormat';                          options = @('set','enable')                      }
        @{ name = 'MailTip';                          options = @('set')                      }
        @{ name = 'MailTipTranslations';                          options = @('set')                      }
        @{ name = 'MailboxContainerGuid';       }
        @{ name = 'MailboxLocations';            }
        @{ name = 'MailboxMoveBatchName';         }
        @{ name = 'MailboxMoveFlags';              }
        @{ name = 'MailboxMoveRemoteHostName';      }
        @{ name = 'MailboxMoveSourceMDB';            }
        @{ name = 'MailboxMoveStatus';                }
        @{ name = 'MailboxMoveTargetMDB';              }
        @{ name = 'MailboxProvisioningConstraint';      }
        @{ name = 'MailboxProvisioningPreferences';      }
        @{ name = 'MailboxRegion';                        }
        @{ name = 'MailboxRegionLastUpdateTime';           }
        @{ name = 'MailboxRelease';                     }
        @{ name = 'MaxReceiveSize';                          options = @('set')                      }
        @{ name = 'MaxSendSize';                          options = @('set')                      }
        @{ name = 'MessageBodyFormat';                          options = @('set')                      }
        @{ name = 'MessageFormat';                          options = @('set','enable')                      }
        @{ name = 'MicrosoftOnlineServicesID';          }
        @{ name = 'ModeratedBy';                          options = @('set')                      }
        @{ name = 'ModerationEnabled';                          options = @('set')                      }
        @{ name = 'Name';                          options = @('set')                      }
        @{ name = 'ObjectCategory';               }
        @{ name = 'ObjectClass';                   }
        @{ name = 'ObjectState';                    }
        @{ name = 'OrganizationId';                  }
        @{ name = 'OrganizationalUnit';               }
        @{ name = 'OriginatingServer';                 }
        @{ name = 'OtherMail';                          }
        @{ name = 'PersistedCapabilities';               }
        @{ name = 'PoliciesExcluded';                     }
        @{ name = 'PoliciesIncluded';                      }
        @{ name = 'PrimarySmtpAddress';                          options = @('default','set','enable')                      }
        @{ name = 'ProhibitSendQuota';                      }
        @{ name = 'ProhibitSendReceiveQuota';                }
        @{ name = 'ProtocolSettings';                        }
        @{ name = 'RecipientLimits';                          options = @('set')                      }
        @{ name = 'RecipientType';                 }
        @{ name = 'RecipientTypeDetails';           }
        @{ name = 'RecoverableItemsQuota';                          options = @('set')                      }
        @{ name = 'RecoverableItemsWarningQuota';                          options = @('set')                      }
        @{ name = 'RejectMessagesFrom';                          options = @('set')                      }
        @{ name = 'RejectMessagesFromDLMembers';                          options = @('set')                      }
        @{ name = 'RejectMessagesFromSendersOrMembers';                          options = @('set')                      }
        @{ name = 'RequireSenderAuthenticationEnabled';                          options = @('set')                      }
        @{ name = 'ResetPasswordOnNextLogon';     }
        @{ name = 'RetainDeletedItemsFor';         }
        @{ name = 'RetentionComment';              }
        @{ name = 'RetentionHoldEnabled';           }
        @{ name = 'RetentionUrl';                    }
        @{ name = 'SKUAssigned';                      }
        @{ name = 'SamAccountName';                          options = @('default','set')                      }
        @{ name = 'SeconaryAddress';                          options = @('set')                      }
        @{ name = 'SendModerationNotifications';                          options = @('set')                      }
        @{ name = 'SimpleDisplayName';                          options = @('set')                      }
        @{ name = 'SingleItemRecoveryEnabled'; }
        @{ name = 'StartDateForRetentionHold';  }
        @{ name = 'StsRefreshTokensValidFrom';   }
        @{ name = 'UMDtmfMap';                          options = @('set')                      }
        @{ name = 'UsageLocation';                }
        @{ name = 'UseMapiRichTextFormat';                          options = @('set')                      }
        @{ name = 'UsePreferMessageFormat';                          options = @('set','enable')                      }
        @{ name = 'UserCertificate';                          options = @('set')                      }
        @{ name = 'UserPrincipalName';                          options = @('default','set')                      }
        @{ name = 'UserSMimeCertificate';                          options = @('set')                      }
        @{ name = 'WhenChanged';                   }
        @{ name = 'WhenChangedUTC';                 }
        @{ name = 'WhenCreated';                     }
        @{ name = 'WhenCreatedUTC';                   }
        @{ name = 'WhenMailboxCreated';                }
        @{ name = 'WhenSoftDeleted';                    }
        @{ name = 'WindowsEmailAddress';                          options = @('set')                      }
        @{ name = 'WindowsLiveID';                       }
        
    )
}


# Default properties and IDM properties are the same
foreach ($class in $Properties.Keys) {
    foreach ($e in $Properties.$class) {
        if (!$e.options) { $e.options = @() }
        if ($e.options.Contains('default')) { $e.options += 'idm' }
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
            $properties = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.CASMailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.CASMailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('enable') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
            $properties = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
            $properties = ($Global:Properties.MailboxDatabase | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.MailboxDatabase | Where-Object { $_.options.Contains('key') }).name
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

function Idm-MailboxPermissionsRead {
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

        Get-ClassMetaData -SystemParams $SystemParams -Class 'MailboxPermission'
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
            $properties = ($Global:Properties.MailboxPermission | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.MailboxPermission | Where-Object { $_.options.Contains('key') }).name
        $properties = @($properties | Where-Object { $_ -ne $key })

        try {
            # https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/get-mailbox?view=exchange-ps
            #
            # Cmdlet availability:
            # v On-premises
            # v Cloud

            LogIO info "Get-MsExchangeMailboxMailbox" -In @call_params
            $mailboxes = Get-MsExchangeMailbox @call_params 
            
            LogIO info "Get-MsExchangeMailboxPermission" -In @call_params
            foreach ($mailbox in $mailboxes) {
                $mailbox | Get-MsExchangeMailboxPermission | Select-Object $properties | Select-Object *,@{label="ExchangeGuid";expression={$mailbox.ExchangeGUID}}
            }
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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('permissionAdd') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
                @{ name = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.Mailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('permissionRemove') } | ForEach-Object {
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

        $key = ($Global:Properties.Mailbox | Where-Object { $_.options.Contains('key') }).name

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
            $properties = ($Global:Properties.MailboxStatistics | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.MailboxStatistics | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('enable') } | ForEach-Object {
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

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name

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
            $properties = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('default') }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name
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
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

                $Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('set') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'optional' }
                }
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

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name

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
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.RemoteMailbox | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.RemoteMailbox | Where-Object { $_.options.Contains('key') }).name

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

function Idm-MailUsersRead {
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

        Get-ClassMetaData -SystemParams $SystemParams -Class 'MailUser'
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
            $properties = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })
        
        try {
            LogIO info "Get-MsExchangeMailUser" -In @call_params
            Get-MsExchangeMailUser @call_params | Select-Object $properties
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}

function Idm-MailUserSet {
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
                @{ name = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.MailUser | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('set') } | ForEach-Object {
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

        $key = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            LogIO info "Set-MsExchangeMailUser" -In @call_params
                $rv = Set-MsExchangeMailUser @call_params
            LogIO info "Set-MsExchangeMailUser" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}

function Idm-MailUserEnable {
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
                @{ name = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.MailUser | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('enable') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'prohibited' }
                }

                $Global:Properties.MailUser | Where-Object { !$_.options.Contains('key') -and $_.options.Contains('enable') } | ForEach-Object {
                    @{ name = $_.name; allowance = 'optional' }
                }
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

        $key = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name

        $call_params = @{
            Identity = $function_params.$key
        }

        if ($system_params.domain_controller.length -gt 0) {
            $call_params.DomainController = $system_params.domain_controller
        }

        $function_params.Remove($key)

        $call_params += $function_params

        try {
            LogIO info "Enable-MsExchangeMailUser" -In @call_params
                $rv = Enable-MsExchangeMailUser @call_params
            LogIO info "Enable-MsExchangeMailUser" -Out $rv

            $rv
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }
    }

    Log info "Done"
}

function Idm-MailUserDisable {
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
                @{ name = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name; allowance = 'mandatory' }

                $Global:Properties.MailUser | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('disable') } | ForEach-Object {
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

        $key = ($Global:Properties.MailUser | Where-Object { $_.options.Contains('key') }).name

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
            LogIO info "Disable-MsExchangeMailUser" -In @call_params
                $rv = Disable-MsExchangeMailUser @call_params
            LogIO info "Disable-MsExchangeMailUser" -Out $rv

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
                            foreach ($opt in $_.options) {
                                if ($opt -notin @('default', 'idm', 'key')) { continue }

                                if ($opt -eq 'idm') {
                                    $opt.Toupper()
                                }
                                else {
                                    $opt.Substring(0,1).Toupper() + $opt.Substring(1)
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
            value = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
    )
}
