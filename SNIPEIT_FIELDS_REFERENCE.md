# Quick Reference: Snipe-IT Custom Fields

## Required Custom Fields for Secure Boot, BIOS Information, and Script Version Tracking

Add these custom fields to your Snipe-IT instance to support Secure Boot, BIOS information, and script version tracking:

**NOTE:** Fields 87, 88, 89, and 90 are used to avoid conflicts with existing custom fields (41-86 are already in use for other purposes like Motherboard, PSU, CPU, RAM, etc.)

### Field 87: Secure Boot Status
```
Field Name: Secure Boot Status
DB Field: _snipeit_secure_boot_status_87
Field Type: Text
Format: Single Line Text
Help Text: Current Secure Boot status (Enabled/Disabled/Not Supported)
Required: No
Fieldsets: Computers
```

**Example Values:**
- `Enabled`
- `Disabled`
- `Not Supported (Legacy BIOS)`
- `Unknown`

---

### Field 88: Secure Boot Certificate Expiry
```
Field Name: Secure Boot Certificate Expiry
DB Field: _snipeit_secure_boot_cert_expiry_88
Field Type: Textarea
Format: Multi-line Text
Help Text: Secure Boot certificate expiration dates with status indicators
Required: No
Fieldsets: Computers
```

**Example Value:**
```
Microsoft Corporation UEFI CA 2011 : 2026-06-30 (EXPIRING SOON, 165 days)
Microsoft Windows Production PCA 2011 : 2026-10-30 (Valid, 287 days)
```

**Status Indicators:**
- `EXPIRED` - Certificate has expired
- `EXPIRING SOON` - Less than 180 days until expiry
- `Valid` - More than 180 days until expiry

---

### Field 89: BIOS Release Date
```
Field Name: BIOS Release Date
DB Field: _snipeit_bios_release_date_89
Field Type: Text (or Date if you prefer)
Format: YYYY-MM-DD
Help Text: Date when the BIOS firmware was released
Required: No
Fieldsets: Computers
```

**Example Value:**
- `2023-05-15`

---

### Field 90: Script Version
```
Field Name: Script Version
DB Field: _snipeit_script_version_90
Field Type: Text
Format: Single Line Text
Help Text: Version of the asset collection script that last reported data
Required: No
Fieldsets: Computers
```

**Example Value:**
- `1.0`

**Version History:**
- `1.0` - Initial version with version reporting

---

### Field 93: Monitors
```
Field Name: Monitors
DB Field: _snipeit_monitors_93
Field Type: Text
Format: Single Line Text
Help Text: Connected monitors (manufacturer code, model, serial)
Required: No
Fieldsets: Computers
```

**Example Value:**
- `DEL DELL U2723QE [ABC123]; SAM S24F350 [H4ZM]`

---

### Field 94: Display Resolution
```
Field Name: Display Resolution
DB Field: _snipeit_display_resolution_94
Field Type: Textarea
Format: Multi-line Text
Help Text: Per display: resolution, Windows scaling, usable space and native mode. LOW-RES / UNDER-1080P flag problems.
Required: No
Fieldsets: Computers
```

Snipe-IT builds the DB field name from the field name and its ID. If the new field is not given ID 94, change `_snipeit_display_resolution_94` in `AssetSelfReport.ps1` to match.

**Example Value:**
```
DEL DELL U2723QE [ABC123]: 1920x1080 @100% = 1920x1080 usable (native 3840x2160) primary LOW-RES
SAM S24F350 [H4ZM]: 1920x1080 @? (native)
```

**Format:** `<monitor>: <resolution> @<scaling> = <usable space> usable (native <mode>) [primary] [flags]`
- `(native)` means the display is at its native mode; `(native ?)` means the monitor did not report one.
- `@?` means scaling is unknown: nobody was signed in, the display is not the primary, or the display layout changed after the user signed in (e.g. they docked). Windows only records an absolute scaling value for the primary display, at sign-in.
- `Display N` is used instead of the monitor name when the display can't be matched to a monitor.
- `Unknown` means the data could not be read; the reason is in the script log.

**Flags:**
- `LOW-RES` - running below the panel's native resolution (e.g. a 4K panel at 1920x1080)
- `UNDER-1080P` - usable space (resolution divided by scaling) is below 1920x1080, the minimum our systems are designed for

Search Snipe-IT for `LOW-RES` or `UNDER-1080P` to list affected machines. The values reflect the newest display layout Windows saved, so a laptop that was docked last will show its docked layout.

---

## How to Add These Fields in Snipe-IT

1. Log in as administrator
2. Go to **Settings** → **Custom Fields**
3. Click **Create New**
4. Fill in the details from above
5. Save each field
6. Add all fields to your "Computers" fieldset (or appropriate fieldset)
7. Associate fieldset with your asset models

## Important Notes

- The field numbers (87, 88, 89, 90) are used in the script to map data to Snipe-IT
- These numbers were chosen to avoid conflicts with existing fields 1-86
- If your Snipe-IT instance uses different field numbers, update the script accordingly
- Certificate information requires administrator privileges to collect
- Legacy BIOS systems will show "Not Supported" for Secure Boot status
- Empty BIOS release dates indicate the BIOS doesn't report this information

## Existing Field Numbers (DO NOT USE)

The following field numbers are already in use (as of this implementation):
- Fields 1-86: Various existing fields (see complete list below)
- Fields 87-90: Secure Boot, BIOS information, and Script Version
- Fields 91-92: Secure Boot certificate expiry date and compliance
- Fields 93-94: Monitors and Display Resolution

### Complete List of Existing Custom Fields (Fields 1-86)

| Field # | Name | DB Field | Help Text | Format | Element | Fieldsets |
|---------|------|----------|-----------|--------|---------|-----------|
| 1 | MAC Address | _snipeit_mac_address_1 | | MAC | text | Asset with MAC Address; Computers; Tablets |
| 2 | CPU | _snipeit_cpu_2 | | ANY | text | Computers |
| 3 | RAM | _snipeit_ram_3 | | ANY | text | Computers |
| 4 | FQDN | _snipeit_fqdn_4 | | ANY | text | Computers |
| 5 | Operating System | _snipeit_operating_system_5 | | ANY | text | Computers |
| 6 | Operating System Build | _snipeit_operating_system_build_6 | | ANY | text | Computers |
| 7 | SKU | _snipeit_sku_7 | | ANY | text | Computers |
| 8 | BIOS Windows License Key | _snipeit_bios_windows_license_key_8 | Windows license key as stored in the BIOS | ANY | text | Computers |
| 9 | IP Address | _snipeit_ip_address_9 | | ANY | text | Computers |
| 10 | Windows Version | _snipeit_windows_version_10 | Release version of Windows | ANY | text | Computers |
| 11 | BIOS | _snipeit_bios_11 | | ANY | text | Computers |
| 12 | Last Reported | _snipeit_last_reported_12 | | ANY | text | Computers |
| 13 | Graphics | _snipeit_graphics_13 | | ANY | text | Computers |
| 14 | Unused 5 | _snipeit_unused_5_14 | | ANY | text | |
| 15 | Boot Drive | _snipeit_boot_drive_15 | | ANY | text | Computers |
| 16 | Internal Media | _snipeit_internal_media_16 | | ANY | textarea | Computers |
| 17 | External Media | _snipeit_external_media_17 | | ANY | text | Computers |
| 18 | Installed Software | _snipeit_installed_software_18 | | ANY | text | Computers |
| 19 | Remote Desktop Users | _snipeit_remote_desktop_users_19 | | ANY | textarea | Computers |
| 20 | RAM Installed | _snipeit_ram_installed_20 | | ANY | text | Computers |
| 21 | Drives | _snipeit_drives_21 | | ANY | text | Computers |
| 22 | Webcam | _snipeit_webcam_22 | | ANY | text | Computers |
| 23 | Applied Updates | _snipeit_applied_updates_23 | | ANY | textarea | Computers |
| 24 | Network Adapters | _snipeit_network_adapters_24 | | ANY | text | Computers |
| 25 | Age | _snipeit_age_25 | | ANY | text | Computers |
| 26 | Last Logged In User | _snipeit_last_logged_in_user_26 | | ANY | text | Computers |
| 27 | Capacity (TB) | _snipeit_capacity_tb_27 | Terabytes as advertised by the manufacturer | NUMERIC | text | Storage |
| 29 | Date Last Formatted | _snipeit_date_last_formatted_29 | | ANY | text | Computers; Tablets |
| 31 | Onsite Warranty Cover | _snipeit_onsite_warranty_cover_31 | | ANY | text | Asset with MAC Address; Computers; Tablets; Storage |
| 32 | Offsite Warranty Cover | _snipeit_offsite_warranty_cover_32 | | ANY | text | Asset with MAC Address; Computers; Tablets; Storage |
| 33 | User Last Formatted | _snipeit_user_last_formatted_33 | | ANY | text | Computers; Tablets |
| 35 | UUID | _snipeit_uuid_35 | | ANY | text | Computers; Tablets |
| 36 | BitLocker Version | _snipeit_bitlocker_version_36 | | ANY | text | Computers; Tablets |
| 38 | BitLocker Summary | _snipeit_bitlocker_summary_38 | | ANY | text | Computers; Tablets |
| 39 | Domain | _snipeit_domain_39 | | ANY | text | Computers; Tablets |
| 40 | Windows UI Language | _snipeit_windows_ui_language_40 | | ANY | text | Computers; Tablets |
| 41 | Motherboard | _snipeit_motherboard_41 | Make and model of motherboard | ANY | listbox | Asset with MAC Address; Computers; Motherboards |
| 42 | PSU | _snipeit_psu_42 | Power Supply Unit make and model | ANY | text | Computers |
| 43 | PSU Wattage | _snipeit_psu_wattage_43 | Max rated wattage of power supply | ANY | text | Computers |
| 44 | CPU Model | _snipeit_cpu_model_44 | | ANY | listbox | Computers |
| 45 | RAM Module Manufacturer | _snipeit_ram_module_manufacturer_45 | Who makes the sticks | ANY | listbox | Computers |
| 46 | RAM Total (GB) | _snipeit_ram_total_gb_46 | Total amount of RAM in GB | NUMERIC | text | Computers |
| 47 | RAM Module Config | _snipeit_ram_module_config_47 | How many sticks and what size each | ANY | listbox | Computers |
| 48 | RAM Speed (MT/s) | _snipeit_ram_speed_mts_48 | Rated speed in megatransfers per second | NUMERIC | text | Computers |
| 49 | RAM CAS Latency (ms) | _snipeit_ram_cas_latency_ms_49 | Column Address Strobe latency in milliseconds | NUMERIC | text | Computers |
| 50 | GPU Manufacturer | _snipeit_gpu_manufacturer_50 | Who makes the GPU chip | ANY | listbox | Computers |
| 51 | Graphics Card AIB | _snipeit_graphics_card_aib_51 | Who makes the graphics card aka add-in board manufacturer | ANY | listbox | Computers |
| 52 | GPU Chipset | _snipeit_gpu_chipset_52 | The important part | ANY | listbox | Computers |
| 53 | GPU VRAM Total (GB) | _snipeit_gpu_vram_total_gb_53 | How big is the frame buffer in gigabytes | NUMERIC | text | Computers |
| 54 | Windows OEM Licence Key | _snipeit_windows_oem_licence_key_54 | OEM key purchased separately and activated with this PC | ANY | text | Computers |
| 55 | Extra Fans | _snipeit_extra_fans_55 | Details of any extra cooling fans | ANY | text | Computers |
| 56 | Extra Peripherals | _snipeit_extra_peripherals_56 | Details of any extra peripherals | ANY | textarea | Computers |
| 57 | Case | _snipeit_case_57 | Chassis housing the components | ANY | listbox | Computers |
| 58 | Misc. Components | _snipeit_misc_components_58 | Any extra bits that go inside the case e.g. expansion cards, CPU mounting & GPU anti sag brackets, sensors etc. | ANY | textarea | Computers |
| 59 | CPU Cooler | _snipeit_cpu_cooler_59 | | ANY | listbox | Computers |
| 60 | Primary Storage Type | _snipeit_primary_storage_type_60 | Boot/system drive(s) | ANY | listbox | Computers; Storage |
| 61 | HDD RPM | _snipeit_hdd_rpm_61 | | NUMERIC | text | Computers; Storage |
| 62 | Primary Storage Interface | _snipeit_primary_storage_interface_62 | How does it connect | ANY | listbox | Computers; Storage |
| 63 | Primary Storage Form Factor | _snipeit_primary_storage_form_factor_63 | Physical size & shape | ANY | listbox | Computers; Storage |
| 64 | Primary Storage Manufacturer | _snipeit_primary_storage_manufacturer_64 | Who makes the drive | ANY | listbox | Computers; Storage |
| 65 | Primary Storage Capacity (TB) | _snipeit_primary_storage_capacity_tb_65 | Terabytes | NUMERIC | text | Computers; Storage |
| 66 | Primary Storage has DRAM Cache | _snipeit_primary_storage_has_dram_cache_66 | Yes is better | ANY | radio | Computers; Storage |
| 67 | Primary Storage Model | _snipeit_primary_storage_model_67 | | ANY | text | Computers; Storage |
| 68 | Motherboard Manufacturer | _snipeit_motherboard_manufacturer_68 | | ANY | listbox | Computers; Motherboards |
| 69 | Motherboard Chipset | _snipeit_motherboard_chipset_69 | | ANY | listbox | Computers; Motherboards |
| 70 | Motherboard Ethernet | _snipeit_motherboard_ethernet_70 | | ANY | listbox | Asset with MAC Address; Computers; Motherboards |
| 71 | Motherboard WiFi | _snipeit_motherboard_wifi_71 | | ANY | listbox | Asset with MAC Address; Computers; Motherboards |
| 72 | Motherboard ECC Support | _snipeit_motherboard_ecc_support_72 | Support for error correction code memory | ANY | radio | Computers; Motherboards |
| 73 | Motherboard No. of Memory Slots | _snipeit_motherboard_no_of_memory_slots_73 | | NUMERIC | text | Computers; Motherboards |
| 74 | CPU Series | _snipeit_cpu_series_74 | | ANY | listbox | Computers |
| 75 | CPU ECC Support | _snipeit_cpu_ecc_support_75 | | ANY | radio | Computers |
| 76 | CPU Microarchitecture | _snipeit_cpu_microarchitecture_76 | | ANY | listbox | Computers |
| 77 | CPU Integrated Graphics | _snipeit_cpu_integrated_graphics_77 | | ANY | listbox | Computers |
| 78 | CPU Core Count | _snipeit_cpu_core_count_78 | | NUMERIC | text | Computers |
| 79 | RAM Timings | _snipeit_ram_timings_79 | | ANY | listbox | Computers |
| 80 | RAM ECC and/or Registered | _snipeit_ram_ecc_andor_registered_80 | | ANY | listbox | Computers |
| 81 | RAM Type | _snipeit_ram_type_81 | | ANY | listbox | Computers |
| 82 | RAM First Word Latency (ms) | _snipeit_ram_first_word_latency_ms_82 | =(CAS Latency (ms) × 2000) / Transfer Rate (MT/s) | NUMERIC | text | Computers |
| 83 | PSU Manufacturer | _snipeit_psu_manufacturer_83 | | ANY | listbox | Computers |
| 84 | PSU Efficiency Rating | _snipeit_psu_efficiency_rating_84 | | ANY | listbox | Computers |
| 85 | PSU EPS/ATX Connectors | _snipeit_psu_epsatx_connectors_85 | | ANY | listbox | Computers |
| 86 | CPU Manufacturer | _snipeit_cpu_manufacturer_86 | Who makes the chip | ANY | listbox | Computers |

**Note:** Some field numbers are missing (e.g., 28, 30, 34, 37) - these may have been deleted or were never created in this instance.

## Verifying the Setup

After creating the fields and deploying the updated script:

1. Run the script on a test machine
2. Check the asset in Snipe-IT
3. Verify all fields are populated with data (including Script Version showing "1.0")
4. Review logs for any warnings about expiring certificates

## Support

For detailed setup instructions, see [SECURE_BOOT_SETUP.md](SECURE_BOOT_SETUP.md)
