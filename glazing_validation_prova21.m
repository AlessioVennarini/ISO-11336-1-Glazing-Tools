close all;
clear; 
clf;
clc;

% --- Load Data ---
% Read the dynamic range coordinates safely using string conversion
ship_start = string(readcell('compl_doc.xlsx', 'Range', 'F3:F3'));
ship_end   = string(readcell('compl_doc.xlsx', 'Range', 'F4:F4'));
cat_start  = string(readcell('compl_doc.xlsx', 'Range', 'G3:G3'));
cat_end    = string(readcell('compl_doc.xlsx', 'Range', 'G4:G4'));
mat_start  = string(readcell('compl_doc.xlsx', 'Range', 'H3:H3'));
mat_end    = string(readcell('compl_doc.xlsx', 'Range', 'H4:H4'));

% Build the range strings dynamically using string concatenation
ship_range = ship_start + ":" + ship_end;
cat_range  = cat_start + ":" + cat_end;
mat_range  = mat_start + ":" + mat_end;

% Load the tables using the dynamic ranges
M_ship = readtable('compl_doc.xlsx', 'Range', ship_range, 'VariableNamingRule', 'preserve');
M_material_cathegory = readtable('compl_doc.xlsx', 'Range', cat_range, 'VariableNamingRule', 'preserve');
M_material = readtable('compl_doc.xlsx', 'Range', mat_range, 'VariableNamingRule', 'preserve');

opts = detectImportOptions('compl_doc.xlsx');
opts.DataRange = 'A61'; 
opts.VariableNamesRange = 'A60'; 
opts.VariableNamingRule = 'preserve';

% Read a temporary table strictly to capture its geometric dimensions
M_temp = readtable('compl_doc.xlsx', opts);
Variable_point = 'A61';
Data_point = 'A60';

% --- Extract Row Numbers into a Structure ---
num_variablepoint = str2double(regexp(Variable_point, '\d+', 'match', 'once'));
num_datapoint     = str2double(regexp(Data_point, '\d+', 'match', 'once'));

% --- Geometric Dimension Approach ---
[num_data_rows, num_columns] = size(M_temp);
final_row = num_datapoint + num_data_rows;

final_column_letter = "";
temp_col = num_columns;
while temp_col > 0
    modulo = mod(temp_col - 1, 26);
    final_column_letter = char(65 + modulo) + final_column_letter;
    temp_col = floor((temp_col - modulo) / 26);
end

calculated_range = sprintf('%s:%s%d', opts.VariableNamesRange, final_column_letter, final_row);
M_glazing = readtable('compl_doc.xlsx', 'Range', calculated_range);

L_H = table2array(M_ship(1,3));
L = table2array(M_ship(2,3));
T_dsw = table2array(M_ship(3,3));
L_int = table2array(M_ship(4,3));
B = table2array(M_ship(5,3));
z_main_deck_dsw = table2array(M_ship(6,3));
z_main_deck_baseline = z_main_deck_dsw + T_dsw;
operational_range = string(M_ship{7,2});
Type_yacht = string(M_ship{8,2});
clean_string_oprange = lower(strtrim(operational_range));
clean_string_typeya = lower(strtrim(Type_yacht));

%% --- Sec. 5.4 Glazed Openings & Skylights: reference waterlines ---
% z_DWL : Design Waterline height above baseline [m] -> used by 5.4(1)(a) "General"
%         and 5.4(2) "Malta Commercial Yacht Code (CYC)" (identical rule).
% z_LL  : Summer Load Line height above baseline [m] -> used by 5.4(5)
%         "International Load Line Convention" (side scuttles only).

% Direct extraction assuming fixed table structure
z_DWL = table2array(M_ship(12, 3));
z_LL  = table2array(M_ship(11, 3));

dwl_missing = isnan(z_DWL);
ll_missing  = isnan(z_LL);

if dwl_missing && ll_missing
    warning('M_ship: Both Design Waterline (z_DWL) and Summer Load Line (z_LL) are missing. All Sec. 5.4 checks will be skipped.');

elseif ~dwl_missing && ll_missing
    warning('M_ship: Design Waterline (z_LL) not found. Sec. 5.4 sill-height check(s) referencing LL will be skipped.');

elseif dwl_missing && ~ll_missing
    warning('M_ship: Design Waterline (z_DWL) not found. Sec. 5.4 sill-height check(s) referencing DWL will be skipped.');
end

% 1b. Sec. 5.4 - Minimum sill freeboard margin above the reference waterline
% Required margin = greater of 2.5% of B or 500 mm (0.5 m)
min_margin_B = 0.025 * B;
min_margin_abs = 0.500;
freeboard_margin_B = max(0.025 * B, 0.5);


%% 1. Define h_std (Global ship variable)
if L <= 75
    h_std = 1.8;
elseif L >= 125
    h_std = 2.3;
else
    h_std = interp1([75, 125], [1.8, 2.3], L);
end

%% 2. Pre-allocate Result Arrays for Glazing Elements
num_glazing = size(M_glazing, 1);

in_critical_zone = false(num_glazing, 1);
valid = false(num_glazing, 1);
Status = repmat("VALID", num_glazing, 1);

% Sec. 5.4 - Glazed Openings & Skylights: sill height above reference waterline
Sill_DWL_limit_vec = nan(num_glazing, 1);   % Min. required sill height above z_DWL [m] - CYC/Malta Code 5.4(1)(a)/(2)
Sill_LL_limit_vec  = nan(num_glazing, 1);   % Min. required sill height above z_LL [m]  - ILLC 5.4(5), side scuttles only
Sill_Height_Status = repmat("VALID", num_glazing, 1);

valid_shutter = false(num_glazing, 1);
Storm_Shutter_Status = repmat("VALID", num_glazing, 1);

valid_deadlight = false(num_glazing, 1);
Deadlight_Status = repmat("VALID", num_glazing, 1);

p_D = zeros(num_glazing, 1);
f_E = zeros(num_glazing, 1); 
p_DE_vec = zeros(num_glazing, 1);

% Fixed ISO table data
L_motor = [24; 30; 40; 50; 60; 70; 80; 90; 100];
p_hull_motor = [70; 70; 70; 70; 76; 84; 91; 98; 108];
p_sup_motor  = [35; 35; 35; 35; 38; 42; 46; 49; 52];
L_sailing = [24; 30; 40; 50; 60; 70];
p_hull_sailing = [70; 70; 70; 83; 96; 109];
p_sup_sailing  = [35; 35; 35; 42; 48; 55];

a_p = zeros(num_glazing, 1);
b_p = zeros(num_glazing, 1);
d = zeros(num_glazing, 1);
t_0_material = zeros(num_glazing, 1);

% Additional pre-allocations to ensure ALL per-glazing calculated variables are saved as workspace vectors
x_vec = zeros(num_glazing, 1);
y_vec = zeros(num_glazing, 1);
h_vec = zeros(num_glazing, 1);
h_baseline_vec = zeros(num_glazing, 1);
h_lowsill_vec = zeros(num_glazing, 1);
Area_vec = zeros(num_glazing, 1);
fi_vec = zeros(num_glazing, 1);
Position_characteristic_vec = strings(num_glazing, 1);
Storm_shutter_opt_vec = strings(num_glazing, 1);
Shape_type_vec = strings(num_glazing, 1);
glazing_type_vec = strings(num_glazing, 1);
macro_category_vec = strings(num_glazing, 1);

% Geometric variables per glazing
b_2_vec = zeros(num_glazing, 1);
delta_z_vec = zeros(num_glazing, 1);
f_h_vec = zeros(num_glazing, 1);
ratio_y_b2_vec = zeros(num_glazing, 1);
f_b_vec = zeros(num_glazing, 1);
f_dir_vec = zeros(num_glazing, 1);
f_PD0_vec = zeros(num_glazing, 1);
p_D0_vec = zeros(num_glazing, 1);
ratio_x_L_vec = zeros(num_glazing, 1);
f_factor_vec = zeros(num_glazing, 1);
b_factor_vec = zeros(num_glazing, 1);
dist_from_side_vec = zeros(num_glazing, 1);
c_factor_vec = zeros(num_glazing, 1);
k_s_vec = zeros(num_glazing, 1);
x1_x_vec = zeros(num_glazing, 1);
x2_x_vec = zeros(num_glazing, 1);
x1_front_base_vec = zeros(num_glazing, 1);
x2_front_base_vec = zeros(num_glazing, 1);
x1_y_vec = zeros(num_glazing, 1);
x2_y_vec = zeros(num_glazing, 1);
x1_vec = zeros(num_glazing, 1);
x2_vec = zeros(num_glazing, 1);
a_factor_vec = zeros(num_glazing, 1);
h_0_vec = zeros(num_glazing, 1);
h_0_std_vec = zeros(num_glazing, 1);
delta_h_0_vec = zeros(num_glazing, 1);
p_D_initial_vec = zeros(num_glazing, 1);
p_D_min_vec = zeros(num_glazing, 1);

% Bending coefficients and geometric factors per glazing
alpha_vec = zeros(num_glazing, 1);
beta_vec = zeros(num_glazing, 1);
aspect_ratio_vec = zeros(num_glazing, 1);

% Ply structures, interlayers, and material properties stored per glazing
plies_t_cell = cell(num_glazing, 1);
plies_mat_cell = cell(num_glazing, 1);
plies_E_cell = cell(num_glazing, 1);
interlayer_t_cell = cell(num_glazing, 1);
G_interlayer_vec = zeros(num_glazing, 1);
E_interlayer_vec = zeros(num_glazing, 1);
E_glass_vec = zeros(num_glazing, 1);
a_dim_vec = zeros(num_glazing, 1);
p_ULS_vec = zeros(num_glazing, 1);
sigma_bending_calc_vec = zeros(num_glazing, 1);
sigma_max_princip_calc_vec = zeros(num_glazing, 1);

% Iteration vector structures for Collaborating plies saved as cell arrays per glazing
h_s_vec_cell          = cell(num_glazing, 1);
t_sk1_vec_cell        = cell(num_glazing, 1);
t_sk2_vec_cell        = cell(num_glazing, 1);
I_s_vec_cell          = cell(num_glazing, 1);
gamma_coef_vec_cell   = cell(num_glazing, 1);
t_eq_w_vec_cell       = cell(num_glazing, 1);
t_1_et_sigma_vec_cell = cell(num_glazing, 1);
t_2_et_sigma_vec_cell = cell(num_glazing, 1);
t_current_f_vec_cell  = cell(num_glazing, 1);

% Deflection variables per glazing
Deflection_Status = strings(num_glazing, 1);
delta_max_val = zeros(num_glazing, 1);
delta_all_val = zeros(num_glazing, 1);
ap_dim_def_vec = zeros(num_glazing, 1);
t_w_vec = zeros(num_glazing, 1);
E_val_vec = zeros(num_glazing, 1);
nu_val_vec = zeros(num_glazing, 1);
M_stiffness_vec = zeros(num_glazing, 1);

% Status & thickness outputs per glazing
Monolithic_Status = strings(num_glazing, 1);
Laminate_Status   = strings(num_glazing, 1);
Glazing_Status    = strings(num_glazing, 1);
t_a_calc          = zeros(num_glazing, 1);
t_commercial_calc = zeros(num_glazing, 1);
t_eq              = zeros(num_glazing, 1);
t_0_panel         = zeros(num_glazing, 1);

% =========================================================================
% STEP 1: MATERIAL STRENGTH PRE-VALIDATION
% =========================================================================
num_materials = size(M_material, 1);
M_material.sigma_C = zeros(num_materials, 1);
M_material.Material_Valid = true(num_materials, 1);
Material_Status_Msg = strings(num_materials, 1);
all_categories = string(M_material_cathegory{:, 2}); 
all_categories_clean = lower(regexprep(strtrim(all_categories), '[^a-zA-Z0-9]', ''));
for m = 1:num_materials
    mat_name = string(M_material{m, 2}); 
    macro_cat_raw = string(M_material{m, 3}); 
    macro_cat_clean = lower(regexprep(strtrim(macro_cat_raw), '[^a-zA-Z0-9]', ''));
    char_strength = M_material{m, 4}; 
    
    idx_cat = find(all_categories_clean == macro_cat_clean, 1);
    
    if isempty(idx_cat)
        warning('Material "%s": Macro-category "%s" not found.', mat_name, macro_cat_raw);
        M_material.Material_Valid(m) = false;
        M_material.sigma_C(m) = NaN;
        Material_Status_Msg(m) = "INVALID (Category Not Found)";
        continue;
    end
    
    min_strength = M_material_cathegory{idx_cat, 5}; 
    max_strength = M_material_cathegory{idx_cat, 6}; 
    
    if char_strength < min_strength
        warning('Material "%s" INVALID: Strength (%g) below minimum (%g).', mat_name, char_strength, min_strength);
        M_material.Material_Valid(m) = false;
        M_material.sigma_C(m) = NaN; 
        Material_Status_Msg(m) = sprintf("INVALID (Strength %.1f < Min %.1f)", char_strength, min_strength);
    else
        if char_strength > max_strength
            warning('Material "%s": Strength (%g) exceeds maximum (%g). Capped.', mat_name, char_strength, max_strength);
            sigma_C = max_strength;
            Material_Status_Msg(m) = sprintf("VALID (Capped at Max %.1f)", max_strength);
        else
            sigma_C = char_strength;
            Material_Status_Msg(m) = "VALID";
        end
        M_material.sigma_C(m) = sigma_C;
    end
end

% --- Convert boolean status into text labels  ---
Validation_Text = string(M_material.Material_Valid);
Validation_Text(Validation_Text == "true") = "Valid";
Validation_Text(Validation_Text == "false") = "Not Valid";

% --- Build Material Structural Validation Report ---
Material_Structural_Validation_Report = table(...
    string(M_material{:, 1}), ...
    string(M_material{:, 2}), ...
    string(M_material{:, 3}), ...
    M_material{:, 4}, ...
    M_material.sigma_C, ...
    Validation_Text, ...
    Material_Status_Msg, ...
    'VariableNames', {'Material_ID', 'Material_Name', 'Macro_Category', ...
                     'Declared_Strength', 'Design_Strength_sigmaC', 'Is_Valid', 'Validation_Status'});

% =========================================================================
% DISPLAY MATERIAL STRUCTURAL VALIDATION REPORT
% =========================================================================

fprintf('\n');
fprintf('=========================================================================\n');
fprintf('                  MATERIAL STRUCTURAL VALIDATION REPORT\n');
fprintf('=========================================================================\n');

for m = 1:height(Material_Structural_Validation_Report)

    fprintf('\n');
    fprintf('-------------------------------------------------------------------------\n');
    fprintf('MATERIAL %s - %s\n', ...
        char(string(Material_Structural_Validation_Report.Material_ID(m))), ...
        char(string(Material_Structural_Validation_Report.Material_Name(m))));
    fprintf('-------------------------------------------------------------------------\n');

    fprintf('  %-32s : %s\n', 'Material ID', ...
        char(string(Material_Structural_Validation_Report.Material_ID(m))));

    fprintf('  %-32s : %s\n', 'Material name', ...
        char(string(Material_Structural_Validation_Report.Material_Name(m))));

    fprintf('  %-32s : %s\n', 'Macro-category', ...
        char(string(Material_Structural_Validation_Report.Macro_Category(m))));

    fprintf('  %-32s : %.4f\n', 'Declared strength', ...
        Material_Structural_Validation_Report.Declared_Strength(m));

    fprintf('  %-32s : %.4f\n', 'Design strength sigma_C', ...
        Material_Structural_Validation_Report.Design_Strength_sigmaC(m));

    fprintf('  %-32s : %s\n', 'Material validity', ...
        char(string(Material_Structural_Validation_Report.Is_Valid(m))));

    fprintf('  %-32s : %s\n', 'Validation status', ...
        char(string(Material_Structural_Validation_Report.Validation_Status(m))));
end

fprintf('\n=========================================================================\n\n');


% =========================================================================
% STEP 2: GLAZING ANALYSIS & PER-PLY \sigma_A VECTORS
% =========================================================================
% Independent working vectors 
gamma_design = zeros(num_glazing, 1);
sigma_A_design = zeros(num_glazing, 1);
sigma_A_plies = cell(num_glazing, 1); % Cell array storing per-ply \sigma_A
ply_mat_indices_cell = cell(num_glazing, 1); % stores material row-indices per glazing
mat_names_clean = lower(regexprep(strtrim(string(M_material{:, 2})), '[^a-zA-Z0-9]', ''));
for j = 1:num_glazing
    ply_mat_indices = [];
    ply_gammas = [];
    
    col_idx = 20; % Starting ply material column 
    while col_idx <= size(M_glazing, 2)
        ply_mat_raw = M_glazing{j, col_idx};
        
        if iscell(ply_mat_raw), ply_mat_str = string(ply_mat_raw{1});
        else, ply_mat_str = string(ply_mat_raw);
        end
        
        if ismissing(ply_mat_str) || strlength(strtrim(ply_mat_str)) == 0 || lower(ply_mat_str) == "nan"
            break;
        end
        
        ply_mat_clean = lower(regexprep(strtrim(ply_mat_str), '[^a-zA-Z0-9]', ''));
        idx_mat = find(mat_names_clean == ply_mat_clean, 1);
        
        if isempty(idx_mat)
            warning('Glazing %d, Col %d: Material "%s" not found in M_material.', j, col_idx, ply_mat_str);
            col_idx = col_idx + 2;
            continue;
        end
        
        macro_cat_raw = string(M_material{idx_mat, 3});
        macro_cat_clean = lower(regexprep(strtrim(macro_cat_raw), '[^a-zA-Z0-9]', ''));
        idx_cat = find(all_categories_clean == macro_cat_clean, 1);
        
        if ~isempty(idx_cat)
            gamma_cat = M_material_cathegory{idx_cat, 7};
            ply_gammas = [ply_gammas; gamma_cat]; %#ok<AGROW>
            ply_mat_indices = [ply_mat_indices; idx_mat]; %#ok<AGROW>
        end
        
        col_idx = col_idx + 2;
    end
    
    if isempty(ply_gammas)
        warning('Glazing row %d: No valid ply materials identified.', j);
        gamma_design(j)   = NaN;
        sigma_A_design(j) = NaN;
        sigma_A_plies{j}  = [];
        ply_mat_indices_cell{j} = [];   
        continue;
    end
    
    gamma_governing = max(ply_gammas);
    gamma_design(j) = gamma_governing;
    
    num_plies = length(ply_mat_indices);
    sigma_A_vec = zeros(1, num_plies);
    
    for k = 1:num_plies
        mat_idx = ply_mat_indices(k);
        sigma_C = M_material.sigma_C(mat_idx);
        sigma_A_vec(k) = sigma_C / gamma_governing;
    end
    
    sigma_A_plies{j} = sigma_A_vec; % Store exact vector for per-ply validation
    ply_mat_indices_cell{j} = ply_mat_indices;   
    sigma_A_design(j) = min(sigma_A_vec); % Baseline value for monolithic t_0
end

%% 3. CALCULATION LOOP FOR EACH GLAZING ELEMENT
for i = 1:num_glazing
    
    x = M_glazing{i, 2};
    y = M_glazing{i, 3};
    h = M_glazing{i, 4};
    h_baseline = h + T_dsw;
    Area = M_glazing{i, 5};    fi = M_glazing{i, 6};
    Position_characteristic = string(M_glazing{i, 7});
    clean_string_poscar = lower(strtrim(Position_characteristic));
    
    Storm_shutter_opt = string(M_glazing{i, 10});
    clean_string_shutter = lower(strtrim(Storm_shutter_opt));
    
    % Store inputs in global vectors
    x_vec(i) = x;
    y_vec(i) = y;
    h_vec(i) = h;
    h_baseline_vec (i) = h_baseline;
    Area_vec(i) = Area;
    fi_vec(i) = fi;
    Position_characteristic_vec(i) = Position_characteristic;
    Storm_shutter_opt_vec(i) = Storm_shutter_opt;

    % --- Geometric Shape Analysis ---
    Shape_type = string(M_glazing{i, 8});
    clean_shape = lower(strtrim(Shape_type));
    Shape_type_vec(i) = Shape_type;
    
    % Safe extraction from cells for inputs with cell2mat handling
    raw_a = M_glazing{i, 13}; if iscell(raw_a), raw_a = cell2mat(raw_a); end; if ischar(raw_a) || isstring(raw_a), raw_a = str2double(raw_a); end
    raw_b = M_glazing{i, 14}; if iscell(raw_b), raw_b = cell2mat(raw_b); end; if ischar(raw_b) || isstring(raw_b), raw_b = str2double(raw_b); end
    raw_d = M_glazing{i, 15}; if iscell(raw_d), raw_d = cell2mat(raw_d); end; if ischar(raw_d) || isstring(raw_d), raw_d = str2double(raw_d); end
    % Developed by Alessio Vennarini (A.V.)
    a_in = raw_a;
    b_in = raw_b;
    d_in = raw_d;
    
    switch clean_shape
        case {'rectangle', 'folded pane', 'quadrangle', 'triangle', 'flat ellipse'}
            macro_category = "rectangle_equiv";
            d(i) = NaN; 
            switch clean_shape
                case {'rectangle', 'folded pane'}, a_p(i) = a_in; b_p(i) = b_in;

                    % --- SWAP FILTER: Enforces a_p >= b_p (ISO Convention) ---
                    if a_p(i) < b_p(i)
                        warning('Glazing %d ("%s"): a_p (%.1f mm) < b_p (%.1f mm). Dimensions swapped to enforce a_p >= b_p.', ...
                            i, char(Shape_type), a_p(i), b_p(i));
                        
                        temp = a_p(i);
                        a_p(i) = b_p(i);
                        b_p(i) = temp;
                    end
                    
                case 'quadrangle', a_p(i) = a_in; b_p(i) = Area / a_p(i);
                case 'triangle', a_p(i) = (2 * a_in) / 3; b_p(i) = (3 * b_in) / 4;
                case 'flat ellipse', a_p(i) = 0.87 * a_in; b_p(i) = 0.87 * b_in;
            end
            
        case {'circle', 'polygon', 'equilateral triangle', 'round ellipse'}
            macro_category = "circle_equiv";
            a_p(i) = NaN; b_p(i) = NaN; 
            switch clean_shape
                case 'circle', d(i) = d_in;
                case 'polygon', d(i) = sqrt(4 / pi) * sqrt(Area);
                case 'equilateral triangle', d(i) = (3 * b_in) / 4;
                case 'round ellipse', d(i) = sqrt(a_in * b_in);
            end
            
        otherwise
            macro_category = "unknown";
            t_0_material(i) = NaN;
            warning('Glazing %d: Unknown geometric shape designation ("%s"). Check shape input (column 8).', i, char(Shape_type));
    end
    macro_category_vec(i) = macro_category;

    h_lowsill = h_baseline - (max([a_p(i) , b_p(i), d(i)])/(2*1000));
    h_lowsill_vec(i) = h_lowsill;

   

    % --- Sec. 5.4 Glazed Openings & Skylights: Sill Height Above Waterline ---
    % (1)(a) General / (2) Malta Commercial Yacht Code (CYC):
    % sill must not be below [Design Waterline + max(2.5%B, 500 mm)]
    Sill_DWL_limit_vec(i) = z_DWL + freeboard_margin_B;
    
    if ~isnan(z_DWL) && h_lowsill < Sill_DWL_limit_vec(i)
        Sill_Height_Status(i) = sprintf("INVALID (Sill h=%.3f m < min %.3f m above Design Waterline - Sec.5.4 CYC/Malta Code)", h_lowsill, Sill_DWL_limit_vec(i));
        warning('Glazing %d: Sill height (%.3f m) is below the minimum required (%.3f m) above the design waterline per Sec. 5.4 (General / Malta Commercial Yacht Code).', i, h_lowsill, Sill_DWL_limit_vec(i));
    end
    
    % (5) International Load Line Convention: applies specifically to side
    % scuttles (glazing on the hull side shell): sill must not be below
    % [Summer Load Line + max(2.5%B, 500 mm)]
    if strcmp(clean_string_poscar, "hull side shell")
        Sill_LL_limit_vec(i) = z_LL + freeboard_margin_B;
        if ~isnan(z_LL) && h_lowsill < Sill_LL_limit_vec(i)
            illc_msg = sprintf("INVALID (Sill h_sill=%.3f m < min %.3f m above Summer Load Line - Sec.5.4 ILLC)", h_lowsill, Sill_LL_limit_vec(i));
            if startsWith(Sill_Height_Status(i), "VALID")
                Sill_Height_Status(i) = illc_msg;
            else
                Sill_Height_Status(i) = Sill_Height_Status(i) + " | " + illc_msg;
            end
            warning('Glazing %d (hull side shell / side scuttle): Sill height (%.3f m) is below the minimum required (%.3f m) above the summer load line per Sec. 5.4 (International Load Line Convention).', i, h_lowsill, Sill_LL_limit_vec(i));
        end
    end

    % --- Storm Shutter Requirement and Robustness Factor f_E ---
    Bulkhead_position = string(M_glazing{i, 11});
    clean_string_bulkhead = lower(strtrim(Bulkhead_position));
    shutter_required = false;
    z_lim_front = 0.02 * L + 2 * h_std;
    z_lim_side= 0.02 * L + h_std;

    switch clean_string_bulkhead
        case "front bulkhead"
            if h < (0.02 * L + 2 * h_std)
                shutter_required = true;
            end
        case "side bulkhead"
            if h < (0.02 * L + h_std)
                shutter_required = true;
            end
        otherwise
            shutter_required = false;
    end

    % Robustness factor f_E
    if shutter_required && clean_string_shutter == "not providing storm shutter and complying with equivalent glazing criteria"
        % Equivalent glazing adopted in lieu of the required storm shutter
        f_E(i) = 1.5;
    else
        % Storm shutter provided, storm shutter not required, or
        % equivalent glazing criteria not satisfied
        f_E(i) = 1.0;
    end
    
    % --- Pressure p_D Calculation ---
    switch clean_string_poscar
        case {"superstructure type a", "deckhouse on or above the freeboard deck", "superstructure on or above the freeboard deck"}
            b_2 = B / 2;
            delta_z = h - z_main_deck_dsw; 
            
            if delta_z <= 1.5, f_h = 1.0;
            elseif delta_z >= 4.0, f_h = 0.0;
            else, f_h = max(0.0, min(1.0, interp1([1.5, 4.0], [1.0, 0.0], delta_z, 'linear', 'extrap')));
            end
            
            ratio_y_b2 = y / b_2;
            if ratio_y_b2 <= 0.85, f_b = 0.0;
            elseif ratio_y_b2 >= 1.0, f_b = 1.0;
            else, f_b = max(0.0, min(1.0, interp1([0.85, 1.0], [0.0, 1.0], ratio_y_b2, 'linear', 'extrap')));
            end
            
            if fi > 90, f_dir = -cos(deg2rad(fi)); else, f_dir = 0.0; end
            
            f_PD0 = f_h * max(f_dir, f_b);
            p_D0  = (1 + f_PD0) * max(15, (12.5 + L/20));
            
            ratio_x_L = x / L;
            f_factor = 0.4284 * (L / 100)^3 - 3.6865 * (L / 100)^2 + 10.951 * (L / 100) - 1.244;
            b_factor = 2.626 * (ratio_x_L)^2 - 2.156 * (ratio_x_L) + 1.436;
            
            dist_from_side = b_2 - y;
            threshold = 0.10 * B;
            if dist_from_side >= threshold, c_factor = 0.85;
            elseif dist_from_side <= 0, c_factor = 1.0; 
            else, c_factor = interp1([0, threshold], [1.0, 0.85], dist_from_side, 'linear');
            end
            
            switch clean_string_oprange
                case "unrestricted range yacht", k_s = 1.00;
                case "intermediate range yacht", k_s = 0.85;
                case "short range yacht", k_s = 0.75;
                otherwise, k_s = 1.0;
            end
            
            if fi < 90
                if ratio_x_L <= 0.25, x1_x = 0.1; x2_x = 0.60;
                elseif ratio_x_L >= 0.75, x1_x = 0.1; x2_x = 0.30;
                else
                    x1_x = interp1([0.25, 0.75], [0.1, 0.1], ratio_x_L, 'linear');
                    x2_x = interp1([0.25, 0.75], [0.60, 0.30], ratio_x_L, 'linear');
                end
            else
                if delta_z <= 1.5, x1_x = 0.83; x2_x = 2.0;
                elseif delta_z >= 6.5, x1_x = 0.66; x2_x = 0.5;
                else
                    x1_x = interp1([1.5, 4.0, 6.5], [0.83, 0.83, 0.66], delta_z, 'linear');
                    x2_x = interp1([1.5, 4.0, 6.5], [2.0, 1.0, 0.5], delta_z, 'linear');
                end
            end
            
            if delta_z <= 1.5, x1_front_base = 0.83; x2_front_base = 2.0;
            elseif delta_z >= 6.5, x1_front_base = 0.66; x2_front_base = 0.5;
            else
                x1_front_base = interp1([1.5, 4.0, 6.5], [0.83, 0.83, 0.66], delta_z, 'linear');
                x2_front_base = interp1([1.5, 4.0, 6.5], [2.0, 1.0, 0.5], delta_z, 'linear');
            end
            
            if ratio_y_b2 <= 0.8, x1_y = 0.1; x2_y = 0.30;
            elseif ratio_y_b2 >= 1.0, x1_y = x1_front_base; x2_y = x2_front_base;
            else
                x1_y = interp1([0.8, 1.0], [0.1, x1_front_base], ratio_y_b2, 'linear');
                x2_y = interp1([0.8, 1.0], [0.30, x2_front_base], ratio_y_b2, 'linear');
            end
            
            x1 = sqrt( (x1_x * cosd(fi))^2 + (x1_y * sind(fi))^2 );
            x2 = sqrt( (x2_x * cosd(fi))^2 + (x2_y * sind(fi))^2 );
            a_factor = x1 * (L / 100) + x2;
            
            
            h_0 = b_factor * f_factor - (p_D0 / (10.05 * a_factor * k_s * c_factor));
            h_0_std = b_factor * f_factor;
            delta_h_0 = p_D0 / (10.05 * a_factor * k_s * c_factor);


            if h < h_0, p_D_initial = 10.05 * a_factor * k_s * (b_factor * f_factor - h) * c_factor;
            else, p_D_initial = p_D0 - 2.27 * (h - h_0);
            end
            
            p_D_min = 3.5 * (1 + f_dir);
            p_D(i) = max(p_D_initial, p_D_min);
            
            % Save calculated pressure sub-variables
            b_2_vec(i) = b_2; delta_z_vec(i) = delta_z; f_h_vec(i) = f_h;
            ratio_y_b2_vec(i) = ratio_y_b2; f_b_vec(i) = f_b; f_dir_vec(i) = f_dir;
            f_PD0_vec(i) = f_PD0; p_D0_vec(i) = p_D0; ratio_x_L_vec(i) = ratio_x_L;
            f_factor_vec(i) = f_factor; b_factor_vec(i) = b_factor; dist_from_side_vec(i) = dist_from_side;
            c_factor_vec(i) = c_factor; k_s_vec(i) = k_s; x1_x_vec(i) = x1_x; x2_x_vec(i) = x2_x;
            x1_front_base_vec(i) = x1_front_base; x2_front_base_vec(i) = x2_front_base;
            x1_y_vec(i) = x1_y; x2_y_vec(i) = x2_y; x1_vec(i) = x1; x2_vec(i) = x2;
            a_factor_vec(i) = a_factor; h_0_vec(i) = h_0; p_D_initial_vec(i) = p_D_initial;
            p_D_min_vec(i) = p_D_min; h_0_std_vec(i) = h_0_std; delta_h_0_vec(i) = delta_h_0;
            
        case {"hull side shell", "superstructure type b"}
            if strcmp(clean_string_poscar, 'hull side shell')
                if contains(clean_string_typeya, 'motor')
                    p_D(i) = interp1(L_motor, p_hull_motor, min(L, 100), 'linear', 'extrap');
                else
                    p_D(i) = interp1(L_sailing, p_hull_sailing, min(L, 70), 'linear', 'extrap');
                end
            else
                if contains(clean_string_typeya, 'motor')
                    p_D(i) = interp1(L_motor, p_sup_motor, min(L, 100), 'linear', 'extrap');
                else
                    p_D(i) = interp1(L_sailing, p_sup_sailing, min(L, 70), 'linear', 'extrap');
                end
            end
            
        otherwise
            p_D(i) = 15;
    end
    
    p_DE_vec(i) = f_E(i) * p_D(i);
    
    % --- Baseline t_0 Determination ---
    sigma_A = sigma_A_design(i);
    
    
    if macro_category == "rectangle_equiv"
        aspect_ratio = max(1.0, min(6.0, a_p(i) / b_p(i)));
        alpha = -4e-5*(aspect_ratio^4) + 0.0008*(aspect_ratio^3) - 0.0058*(aspect_ratio^2) + 0.0184*aspect_ratio - 0.0094;
        beta  = -0.0028*(aspect_ratio^4) + 0.0503*(aspect_ratio^3) - 0.342*(aspect_ratio^2) + 1.0397*aspect_ratio - 0.4592;
        t_0_material(i) = b_p(i) * sqrt((beta * p_DE_vec(i)) / (1000 * sigma_A));
        aspect_ratio_vec(i) = aspect_ratio;
        alpha_vec(i) = alpha;
        beta_vec(i) = beta;
        
    elseif macro_category == "circle_equiv"
        t_0_material(i) = 0.5 * d(i) * sqrt((1.21 * p_DE_vec(i)) / (1000 * sigma_A));
    end
    

    % =========================================================================
    % GLAZING CONFIGURATION ROUTING VIA SWITCH-CASE (ISO 11336-1)
    % =========================================================================

    % Extract all ply thicknesses dynamically for this glazing row once
    ply_thicknesses_vec = [];
    col_idx = 20; % Starting ply material column
    while col_idx <= size(M_glazing, 2)
        mat_check = M_glazing{i, col_idx};
        if iscell(mat_check), mat_check = mat_check{1}; end
        
        % Force scalar conversion safely
        if isempty(mat_check)
            break;
        end
        mat_str = strtrim(string(mat_check(1)));
        
        if ismissing(mat_str) || strlength(mat_str) == 0 || strcmpi(mat_str, "nan")
            break;
        end
        
        % Assuming thickness is stored in the column immediately following the material name column
        thick_val = M_glazing{i, col_idx + 1};
        if iscell(thick_val), thick_val = cell2mat(thick_val); end
        if ~isempty(thick_val) && ~isnan(double(thick_val(1)))
            ply_thicknesses_vec = [ply_thicknesses_vec; double(thick_val(1))];
        end
        col_idx = col_idx + 2; % Jump to next ply pair (material + thickness)
    end

    num_tot_cols = size(M_glazing, 2);
    
    % Robust string cleaning: removes hidden tabs, newlines, and strips leading/trailing spaces
    raw_glazing_type = string(M_glazing{i, 16});
    glazing_type = strtrim(raw_glazing_type);
    glazing_type = regexprep(glazing_type, '\s+', ' '); 
    glazing_type_vec(i) = glazing_type;
    
    % Initialize status outputs and equivalent thickness variables for panel i
    Monolithic_Status(i) = "VALID"; 
    Laminate_Status(i)   = "VALID"; 
    t_a_calc(i)          = NaN; 
    t_commercial_calc(i) = NaN; 
    t_eq(i)              = NaN; 
    
    % Retrieve per-ply data calculated in Step 2
    current_sigma_A_vec = sigma_A_plies{i};
    if exist('plies_validity', 'var') && length(plies_validity) >= i
        current_ply_validity = plies_validity{i};
    else
        current_ply_validity = ones(length(current_sigma_A_vec), 1);
    end

    % Global t_0 calculation valid for the panel with NaN debugging
    sig_A_gov = min(current_sigma_A_vec(:));
    val_pDE = p_DE_vec(i); if iscell(val_pDE), val_pDE = cell2mat(val_pDE); end; val_pDE = double(val_pDE(1));
    val_bp  = b_p(i);  if iscell(val_bp),  val_bp  = cell2mat(val_bp);  end; val_bp  = double(val_bp(1));
    val_d   = d(i);    if iscell(val_d),   val_d   = cell2mat(val_d);    end; val_d   = double(val_d(1));
    
    % Safe extraction and check for beta
    if exist('beta', 'var') && ~isempty(beta)
        val_beta = beta;
    else
        val_beta = 1; % Safe default fallback value
    end
    
    if iscell(val_beta), val_beta = cell2mat(val_beta); end
    
    if isempty(val_beta) || (isnumeric(val_beta) && isnan(val_beta(1)))
        val_beta = 1; 
    end
    val_beta = double(val_beta(1));

    if strcmp(macro_category, "rectangle_equiv")
        if isnan(val_bp)
            warning('Glazing %d: b_p is NaN for shape "%s". Check geometry inputs (columns 13-15).', i, clean_shape);
        end
        t_0_panel(i) = val_bp * sqrt((val_beta * val_pDE) ./ (1000 * sig_A_gov));
    elseif strcmp(macro_category, "circle_equiv")
        if isnan(val_d)
            warning('Glazing %d: d is NaN for shape "%s". Check geometry inputs (columns 12-14).', i, clean_shape);
        end
        t_0_panel(i) = 0.5 * val_d * sqrt((1.21 * val_pDE) ./ (1000 * sig_A_gov));
    else
        t_0_panel(i) = NaN;
        warning('Glazing %d: Macro-category is unknown ("%s"). Unable to compute t_0_panel(i).', i, char(macro_category));
    end

    switch glazing_type
    
        case "Monolithic"
            % =================================================================
            % 1. MONOLITHIC PANEL RULES & COHERENCE CHECKS
            % =================================================================
            current_status_monolithic = "VALID";
            
            Wheelhouse_pos = lower(strtrim(string(M_glazing{i, 12})));
            t_min = M_glazing{i, 17}; 
            t_a = M_glazing{i, 21}; % Actual manual thickness entered for monolithic (col 21)
            
            mat_val_mono = M_glazing{i, 20}; % Material of the monolithic panel (col 20)
            if iscell(mat_val_mono), clean_mat_mono = lower(strtrim(string(mat_val_mono{1}))); 
            else 
                clean_mat_mono = lower(strtrim(string(mat_val_mono))); 
            end
            
            mat_match_idx = find(lower(strtrim(string(M_material{:, 2}))) == clean_mat_mono, 1);
            if ~isempty(mat_match_idx)
                mat_type = lower(strtrim(string(M_material{mat_match_idx, 3})));
                safety_char = lower(strtrim(string(M_material{mat_match_idx, 5})));
            else
                mat_type = ""; safety_char = "";
            end
            
            % --- Wheelhouse Position Constraint ---
            if current_status_monolithic == "VALID" && contains(mat_type, 'glass') && ...
                 (contains(Wheelhouse_pos, 'in weelhouse, front position') || ...
                  contains(Wheelhouse_pos, 'in weelhouse, side position'))
                current_status_monolithic = "INVALID (Laminated safety glass required in front/side wheelhouse)";
                warning('Glazing %d (Monolithic): Laminated safety glass is required in front/side wheelhouse positions.', i);
            end
            
            % --- t_0_panel Availability Check ---
            if current_status_monolithic == "VALID" && isnan(t_0_panel(i))
                current_status_monolithic = "INVALID (t_0_panel is NaN)";
                warning('Glazing %d (Monolithic): t_0_panel is NaN, skipping thickness validation.', i);
            end
            % --- Fragmentation Safety Check ---
            if current_status_monolithic == "VALID" && contains(mat_type, 'toughened') && ~contains(safety_char, 'present')
                current_status_monolithic = "INVALID (Toughened glass lacks fragmentation characteristic)";
                warning('Glazing %d (Monolithic): Toughened glass lacks the required fragmentation characteristic.', i);
            end
            
            % --- Theoretical and Commercial Design Calculations ---
            t_a_calc(i) = max(t_0_panel(i), t_min);
            t_commercial_calc(i) = ceil(t_a_calc(i));
            t_eq(i) = t_a_calc(i);
            
            % --- Thickness Check ---
            if current_status_monolithic == "VALID" && t_a < t_commercial_calc(i)
                current_status_monolithic = sprintf("INVALID (Thickness too low: required %d mm, actual %g mm)", t_commercial_calc(i), t_a);
                warning('Glazing %d (Monolithic): Actual thickness (%.2f mm) is below required commercial thickness (%d mm).', i, t_a, t_commercial_calc(i));
            end
            
            % --- Monolithic Layout Rule: Check that EVERYTHING is empty from column 22 onwards ---
            if current_status_monolithic == "VALID" && num_tot_cols >= 22
                sub_table_destra = M_glazing(i, 22:end);
                has_extra = false;
                for col_var = 1:width(sub_table_destra)
                    val_cell = sub_table_destra{1, col_var};
                    if iscell(val_cell)
                        if ~isempty(val_cell{1}) && ~any(isnan(val_cell{1})) && ~strcmpi(string(val_cell{1}), "nan")
                            has_extra = true; break;
                        end
                    else
                        if ~isempty(val_cell) && ~any(isnan(val_cell)) && ~strcmpi(string(val_cell), "nan")
                            has_extra = true; break;
                        end
                    end
                end
                
                if has_extra
                    warning('Glazing %d: Data detected past column 21. The panel is set as Monolithic but should only have 1 ply.', i);
                    current_status_monolithic = "INVALID (Data detected past column 21 for Monolithic profile)";
                end
            end
            
            % Write definitive status for monolithic
            Monolithic_Status(i) = current_status_monolithic;
            
        case "Type A laminate with indipendent plies"
            % =================================================================
            % 2. TYPE A LAMINATE WITH INDEPENDENT PLIES
            % =================================================================
            current_status_typeA_indipendent = "VALID";
            plies_t = []; plies_mat = string.empty;
            col_idx = 20; 
            
            while col_idx <= num_tot_cols
                mat_val = M_glazing{i, col_idx};
                if iscell(mat_val), str_mat = string(mat_val{1}); else, str_mat = string(mat_val); end
                if ismissing(str_mat) || strtrim(str_mat) == "" || lower(strtrim(str_mat)) == "nan", break; end
                
                if col_idx + 1 <= num_tot_cols
                    t_val = M_glazing{i, col_idx + 1};
                    if iscell(t_val), val_t = t_val{1}; else, val_t = t_val; end
                    if isnan(val_t), break; end
                else
                    break; 
                end
                plies_mat = [plies_mat; lower(strtrim(str_mat))]; %#ok<AGROW>
                plies_t = [plies_t; val_t]; %#ok<AGROW>
                col_idx = col_idx + 2;
            end
            
            plies_t_cell{i} = plies_t;
            plies_mat_cell{i} = plies_mat;
            
            n_plies = length(plies_t);
            is_homogeneous = (n_plies <= 1) || all(plies_mat == plies_mat(1));
            
            invalid_ply_idx = find(current_ply_validity(1:n_plies) == 0, 1);

             % --- t_0_panel Availability Check ---
            if current_status_typeA_indipendent == "VALID" && isnan(t_0_panel(i))
                current_status_typeA_indipendent = "INVALID (t_0_panel is NaN)";
                warning('Glazing %d ("Type A laminate with indipendent plies"): t_0_panel is NaN, skipping thickness validation.', i);
            end

            if current_status_typeA_indipendent == "VALID" && ~isempty(invalid_ply_idx)
                bad_mat = plies_mat(invalid_ply_idx);
                current_status_typeA_indipendent = sprintf("INVALID (Ply %d material '%s' has validity 0)", invalid_ply_idx, bad_mat);
                warning('Glazing %d (Laminate Type A): Ply %d (%s) has material validity 0.', i, invalid_ply_idx, bad_mat);
                t_eq(i) = NaN;
            elseif current_status_typeA_indipendent == "VALID" && ~is_homogeneous
                warning('Glazing %d: Material mismatch in Type A Independent. Intercepting and rerouting inline to Type B processing.', i);
                % --- INLINE TYPE B REROUTING BLOCK ---
                E_col_cat = table2array(M_material_cathegory(:, 8));
                valid_E_values = E_col_cat(~isnan(E_col_cat) & E_col_cat > 0);
                if isempty(valid_E_values)
                    error('Database Error: No valid Young''s modulus values (E > 0) found in M_material_cathegory (Column 8).');
                end
                E_fallback = max(valid_E_values);
                
                plies_E = zeros(n_plies, 1);
                for j = 1:n_plies
                    mat_match_idx = find(lower(strtrim(string(M_material{:, 2}))) == plies_mat(j), 1);
                    if isempty(mat_match_idx)
                        plies_E(j) = E_fallback; 
                    else
                        macro_cat = lower(strtrim(string(M_material{mat_match_idx, 3})));
                        cat_match_idx = find(all_categories_clean == lower(regexprep(macro_cat, '[^a-zA-Z0-9]', '')), 1);
                        if isempty(cat_match_idx), plies_E(j) = E_fallback; else, plies_E(j) = M_material_cathegory{cat_match_idx, 8}; end
                    end
                    if isnan(plies_E(j)) || plies_E(j) <= 0, plies_E(j) = E_fallback; end
                end
                plies_E_cell{i} = plies_E;
                
                sum_E_t3 = sum(plies_E .* (plies_t.^3));
                [~, min_idx_b] = min(plies_t);
                t_eq(i) = sqrt(sum_E_t3 / (plies_E(min_idx_b) * plies_t(min_idx_b)));
                
                if t_eq(i) < t_0_panel(i)
                    current_status_typeA_indipendent = sprintf("INVALID (Type B t_eq = %.2f mm < required t_0 = %.2f mm)", t_eq(i), t_0_panel(i));
                    warning('Glazing %d (Laminate Type A - Rerouted to Type B): Equivalent thickness (%.2f mm) is below required t_0 (%.2f mm).', i, t_eq(i), t_0_panel(i));
                end
                
            elseif current_status_typeA_indipendent == "VALID" && n_plies < 2
                warning('Glazing %d: Less than 2 valid plies parsed for Independent profile.', i);
                current_status_typeA_indipendent = "INVALID (Insufficient plies)";
            else
                if current_status_typeA_indipendent == "VALID"
                    sum_cubes = sum(plies_t.^3);
                    t_eq_j = sqrt(sum_cubes ./ plies_t);
                    [t_eq(i), ~] = min(t_eq_j);
                    
                    if t_eq(i) < t_0_panel(i)
                        current_status_typeA_indipendent = sprintf("INVALID (t_eq = %.2f mm < required t_0 = %.2f mm)", t_eq(i), t_0_panel(i));
                        warning('Glazing %d (Laminate Type A - Independent): Equivalent thickness (%.2f mm) is below required t_0 (%.2f mm).', i, t_eq(i), t_0_panel(i));
                    end
                end
            end
            
            Laminate_Status(i) = current_status_typeA_indipendent;
            
        case "Type A laminate with collaborating plies"
            % =================================================================
            % 3. TYPE A LAMINATE WITH COLLABORATING PLIES
            % =================================================================
            current_status_typeA_collaborating = "VALID";
            t_min = M_glazing{i, 17};
            plies_t = []; plies_mat = string.empty; interlayer_t = [];
            col_idx = 20; 
            
            while col_idx <= num_tot_cols
                mat_val = M_glazing{i, col_idx};
                
                if isempty(mat_val)
                    break;
                end
                
                if iscell(mat_val), str_mat = string(mat_val{1}); else, str_mat = string(mat_val); end
                if ismissing(str_mat) || strtrim(str_mat) == "" || lower(strtrim(str_mat)) == "nan"
                    break; 
                end
                
                % Check glass thickness
                if col_idx + 1 <= num_tot_cols
                    t_val = M_glazing{i, col_idx + 1};
                    if isempty(t_val)
                        break;
                    end
                    if iscell(t_val), val_t = t_val{1}; else, val_t = t_val; end
                    if isnan(double(val_t)), break; end
                else
                    break; 
                end
                
                plies_mat = [plies_mat; lower(strtrim(str_mat))]; %#ok<AGROW>
                plies_t = [plies_t; double(val_t)]; %#ok<AGROW>
                
                col_idx = col_idx + 2;
            end
            
            % --- Generate identical interlayers from the single t_int in column 17 ---
            n_plies = length(plies_t);
            
            raw_tint = M_glazing{i, 17};
            
            if iscell(raw_tint)
                raw_tint = raw_tint{1};
            end
            
            if isempty(raw_tint) || ...
               (isnumeric(raw_tint) && isnan(raw_tint))
            
                interlayer_t = [];
            
            else
                t_int_value = double(raw_tint(1));
            
                % One interlayer between each pair of adjacent plies
                interlayer_t = repmat(t_int_value, max(n_plies - 1, 0), 1);
            end
            
            plies_t_cell{i} = plies_t;
            plies_mat_cell{i} = plies_mat;
            interlayer_t_cell{i} = interlayer_t;
            
            n_plies = length(plies_t);
            is_homogeneous = (n_plies <= 1) || all(plies_mat == plies_mat(1));
            invalid_ply_idx = find(current_ply_validity(1:n_plies) == 0, 1);

            % --- t_0_panel Availability Check ---
            if current_status_typeA_collaborating == "VALID" && isnan(t_0_panel(i))
                current_status_typeA_collaborating = "INVALID (t_0_panel is NaN)";
                warning('Glazing %d ("Type A laminate with collaborating plies"): t_0_panel is NaN, skipping thickness validation.', i);
            end

            if current_status_typeA_collaborating == "VALID" && ~isempty(invalid_ply_idx)
                bad_mat = plies_mat(invalid_ply_idx);
                current_status_typeA_collaborating = sprintf("INVALID (Ply %d material '%s' has validity 0)", invalid_ply_idx, bad_mat);
                warning('Glazing %d (Laminate Type A Collaborating): Ply %d (%s) has material validity 0.', i, invalid_ply_idx, bad_mat);
                t_eq(i) = NaN;
            elseif current_status_typeA_collaborating == "VALID" && ~is_homogeneous
                warning('Glazing %d: Material mismatch in Type A Collaborating. Intercepting and rerouting inline to Type B processing.', i);
                % --- INLINE TYPE B REROUTING BLOCK ---
                E_col_cat = table2array(M_material_cathegory(:, 8));
                valid_E_values = E_col_cat(~isnan(E_col_cat) & E_col_cat > 0);
                if isempty(valid_E_values)
                    error('Database Error: No valid Young''s modulus values (E > 0) found in M_material_cathegory (Column 8).');
                end
                E_fallback = max(valid_E_values);
                
                plies_E = zeros(n_plies, 1);
                for j = 1:n_plies
                    mat_match_idx = find(lower(strtrim(string(M_material{:, 2}))) == plies_mat(j), 1);
                    if isempty(mat_match_idx)
                        plies_E(j) = E_fallback; 
                    else
                        macro_cat = lower(strtrim(string(M_material{mat_match_idx, 3})));
                        cat_match_idx = find(all_categories_clean == lower(regexprep(macro_cat, '[^a-zA-Z0-9]', '')), 1);
                        if isempty(cat_match_idx), plies_E(j) = E_fallback; else, plies_E(j) = M_material_cathegory{cat_match_idx, 8}; end
                    end
                    if isnan(plies_E(j)) || plies_E(j) <= 0, plies_E(j) = E_fallback; end
                end
                plies_E_cell{i} = plies_E;
                
                sum_E_t3 = sum(plies_E .* (plies_t.^3));
                [~, min_idx_b] = min(plies_t);
                t_eq(i) = sqrt(sum_E_t3 / (plies_E(min_idx_b) * plies_t(min_idx_b)));
                
                if t_eq(i) < t_0_panel(i)
                    current_status_typeA_collaborating = sprintf("INVALID (Type B t_eq = %.2f mm < required t_0 = %.2f mm)", t_eq(i), t_0_panel(i));
                    warning('Glazing %d (Laminate Type A Collaborating - Rerouted to Type B): Equivalent thickness (%.2f mm) is below required t_0 (%.2f mm).', i, t_eq(i), t_0_panel(i));
                end
                
            elseif current_status_typeA_collaborating == "VALID" && n_plies < 2
                warning('Glazing %d: Less than 2 valid plies parsed for Collaborating profile.', i);
                current_status_typeA_collaborating = "INVALID (Insufficient plies)";
            else
                if current_status_typeA_collaborating == "VALID"
                    G_interlayer = NaN; val_col18 = M_glazing{i, 18}; val_col19 = M_glazing{i, 19};
                    if iscell(val_col18), is_nan18 = isempty(val_col18{1}) || any(isnan(val_col18{1})); else, is_nan18 = isnan(val_col18); end
                    if iscell(val_col19), is_nan19 = isempty(val_col19{1}) || any(isnan(val_col19{1})); else, is_nan19 = isnan(val_col19); end
                    if ~is_nan18
                        if iscell(val_col18), G_interlayer = val_col18{1}; else, G_interlayer = val_col18; end
                    elseif ~is_nan19
                        if iscell(val_col19), E_interlayer = val_col19{1}; else, E_interlayer = val_col19; end
                        G_interlayer = E_interlayer / 3;
                        E_interlayer_vec(i) = E_interlayer;
                    else
                        warning('Glazing %d: Interlayer columns are empty. Defaulting to independent behavior.', i);
                    end
                    G_interlayer_vec(i) = G_interlayer;
                    
                    if macro_category == "rectangle_equiv"
                        a_dim = min(a_p(i), b_p(i));
                    elseif macro_category == "circle_equiv"
                        a_dim = d(i);
                    end

                    if isnan(a_dim), a_dim = b_in; end
                    a_dim_vec(i) = a_dim;
                    
                    E_col_cat = table2array(M_material_cathegory(:, 8));
                    valid_E_values = E_col_cat(~isnan(E_col_cat) & E_col_cat > 0);
                    if isempty(valid_E_values)
                        error('Database Error: No valid Young''s modulus values (E > 0) found in M_material_cathegory (Column 8).');
                    end
                    E_fallback = max(valid_E_values);
                    
                    mat_match_idx = find(lower(strtrim(string(M_material{:, 2}))) == plies_mat(1), 1);
                    if ~isempty(mat_match_idx)
                        cat_name = lower(strtrim(string(M_material{mat_match_idx, 3})));
                        cat_idx = find(all_categories_clean == lower(regexprep(cat_name, '[^a-zA-Z0-9]', '')), 1);
                        if ~isempty(cat_idx), E_glass = M_material_cathegory{cat_idx, 8}; else, E_glass = E_fallback; end
                    else
                        E_glass = E_fallback;
                    end
                    E_glass_vec(i) = E_glass;
                    
                % --- Input Safety Check ---
                if any(isnan(plies_t)) || any(isnan(interlayer_t)) || isnan(E_glass) || isnan(G_interlayer) || isnan(a_dim)
                    warning('Error: NaN values detected in input vectors or constants: plies_t, interlayer_t, E_glass, G_interlayer, a_dim. Please check your upstream data.');
                end

                iter_mode = 'dynamic'; 
                
                % Initialization
                num_iterations = n_plies - 1;
                h_s_vec          = zeros(num_iterations, 1);
                t_sk1_vec        = zeros(num_iterations, 1);
                t_sk2_vec        = zeros(num_iterations, 1);
                I_s_vec          = zeros(num_iterations, 1);
                gamma_coef_vec   = zeros(num_iterations, 1);
                t_eq_w_vec       = zeros(num_iterations, 1);
                t_1_et_sigma_vec = zeros(num_iterations, 1);
                t_2_et_sigma_vec = zeros(num_iterations, 1);
                t_current_f_vec  = zeros(num_iterations, 1);
                
                t_current_accumulated = plies_t(1); 

                for k = 1:num_iterations
                    t_current_1 = t_current_accumulated;        
                    t_current_2 = plies_t(k + 1);               
                    t_int = interlayer_t(k);
                    
                    t_base_1 = plies_t(k); 
                    t_base_2 = plies_t(k + 1);
                    
                    switch lower(iter_mode)
                        case 'dynamic'
                            h_s = 0.5 * (t_current_1 + t_current_2) + t_int;
                            t_sk2 = (h_s * t_current_2) / (t_current_1 + t_current_2);
                            t_sk1 = (h_s * t_current_1) / (t_current_1 + t_current_2);
                            I_s = t_current_1 * (t_sk2^2) + t_current_2 * (t_sk1^2);
                            
                        case 'fix_all'
                            h_s = 0.5 * (t_base_1 + t_base_2) + t_int;
                            t_sk2 = (h_s * t_base_2) / (t_base_1 + t_base_2);
                            t_sk1 = (h_s * t_base_1) / (t_base_1 + t_base_2);
                            I_s = t_base_1 * (t_sk2^2) + t_base_2 * (t_sk1^2);
                            
                        case 'fix_hs'
                            h_s = 0.5 * (t_base_1 + t_base_2) + t_int;
                            t_sk2 = (h_s * t_base_2) / (t_base_1 + t_base_2);
                            t_sk1 = (h_s * t_base_1) / (t_base_1 + t_base_2);
                            I_s = t_current_1 * (t_sk2^2) + t_current_2 * (t_sk1^2);
                            
                        case 'fix_is'
                            h_s_base = 0.5 * (t_base_1 + t_base_2) + t_int;
                            t_sk2_base = (h_s_base * t_base_2) / (t_base_1 + t_base_2);
                            t_sk1_base = (h_s_base * t_base_1) / (t_base_1 + t_base_2);
                            I_s = t_base_1 * (t_sk2_base^2) + t_base_2 * (t_sk1_base^2);
                            
                            h_s = 0.5 * (t_current_1 + t_current_2) + t_int;
                            t_sk2 = (h_s * t_current_2) / (t_current_1 + t_current_2);
                            t_sk1 = (h_s * t_current_1) / (t_current_1 + t_current_2);
                            
                        otherwise
                            error('Invalid iter_mode specified.');
                    end
                    
                    if ~isnan(G_interlayer) && G_interlayer > 0
                        gamma_coef = 1 / (1 + 9.6 * (E_glass / G_interlayer) * (I_s * t_int) / ((h_s^2) * (a_dim^2)));
                    else
                        gamma_coef = 0;
                    end

                    t_eq_w = ((t_current_1^3) + (t_current_2^3) + 12 * gamma_coef * I_s)^(1/3);
                    t_1_et_sigma = sqrt((t_eq_w^3) / (t_current_1 + 2 * gamma_coef * t_sk2));
                    t_2_et_sigma = sqrt((t_eq_w^3) / (t_current_2 + 2 * gamma_coef * t_sk1));
                    
                    t_current_f = min(t_1_et_sigma, t_2_et_sigma);
                    
                    h_s_vec(k)          = h_s;
                    t_sk1_vec(k)        = t_sk1;
                    t_sk2_vec(k)        = t_sk2;
                    I_s_vec(k)          = I_s;
                    gamma_coef_vec(k)   = gamma_coef;
                    t_eq_w_vec(k)       = t_eq_w;
                    t_1_et_sigma_vec(k) = t_1_et_sigma;
                    t_2_et_sigma_vec(k) = t_2_et_sigma;
                    t_current_f_vec(k)  = t_current_f;
                    
                    t_current_accumulated = t_current_f; 
                end
                
                % Store iteration vector structures per glazing in workspace cell arrays
                h_s_vec_cell{i}          = h_s_vec;
                t_sk1_vec_cell{i}        = t_sk1_vec;
                t_sk2_vec_cell{i}        = t_sk2_vec;
                I_s_vec_cell{i}          = I_s_vec;
                gamma_coef_vec_cell{i}   = gamma_coef_vec;
                t_eq_w_vec_cell{i}       = t_eq_w_vec;
                t_1_et_sigma_vec_cell{i} = t_1_et_sigma_vec;
                t_2_et_sigma_vec_cell{i} = t_2_et_sigma_vec;
                t_current_f_vec_cell{i}  = t_current_f_vec;
                
                t_eq(i) = t_current_accumulated;
                
                if t_eq(i) < t_0_panel(i)
                    current_status_typeA_collaborating = sprintf("INVALID (t_eq = %.2f mm < required t_0 = %.2f mm)", t_eq(i), t_0_panel(i));
                    warning('Glazing %d (Laminate Type A - Collaborating): Equivalent thickness t_eq (%.2f mm) is below required t_0 (%.2f mm).', i, t_eq(i), t_0_panel(i));
                end
            end
            end
            
            Laminate_Status(i) = current_status_typeA_collaborating;
            
        case "Type B laminate"
            % =================================================================
            % 4. TYPE B LAMINATE (DIFFERENT MATERIALS)
            % =================================================================
            current_status_typeB = "VALID";
            plies_t = []; plies_mat = string.empty;
            col_idx = 20; 
            
            while col_idx <= num_tot_cols
                mat_val = M_glazing{i, col_idx};
                if iscell(mat_val), str_mat = string(mat_val{1}); else, str_mat = string(mat_val); end
                if ismissing(str_mat) || strtrim(str_mat) == "" || lower(strtrim(str_mat)) == "nan", break; end
                
                if col_idx + 1 <= num_tot_cols
                    t_val = M_glazing{i, col_idx + 1};
                    if iscell(t_val), val_t = t_val{1}; else, val_t = t_val; end
                    if isnan(val_t), break; end
                else
                    break;
                end
                plies_mat = [plies_mat; lower(strtrim(str_mat))]; %#ok<AGROW>
                plies_t = [plies_t; val_t]; %#ok<AGROW>
                col_idx = col_idx + 2;
            end
            
            plies_t_cell{i} = plies_t;
            plies_mat_cell{i} = plies_mat;
            
            n_plies = length(plies_t);
            invalid_ply_idx = find(current_ply_validity(1:n_plies) == 0, 1);
            
            % --- t_0_panel Availability Check ---
            if current_status_typeB == "VALID" && isnan(t_0_panel(i))
                current_status_typeB = "INVALID (t_0_panel is NaN)";
                warning('Glazing %d ("Type B laminate"): t_0_panel is NaN, skipping thickness validation.', i);
            end
            
            if current_status_typeB == "VALID" && ~isempty(invalid_ply_idx)
                bad_mat = plies_mat(invalid_ply_idx);
                current_status_typeB = sprintf("INVALID (Type B Ply %d material '%s' has validity 0)", invalid_ply_idx, bad_mat);
                warning('Glazing %d (Laminate Type B): Ply %d (%s) has material validity 0.', i, invalid_ply_idx, bad_mat);
                t_eq(i) = NaN;
            elseif current_status_typeB == "VALID" && n_plies < 2
                warning('Glazing %d: Less than 2 valid plies processed for Type B profile.', i);
                t_eq(i) = NaN;
                current_status_typeB = "INVALID (Insufficient Type B plies)";
            else
                if current_status_typeB == "VALID"
                    E_col_cat = table2array(M_material_cathegory(:, 8));
                    valid_E_values = E_col_cat(~isnan(E_col_cat) & E_col_cat > 0);
                    
                    if isempty(valid_E_values)
                        error('Database Error: No valid Young''s modulus values (E > 0) found in M_material_cathegory (Column 8).');
                    end
                    
                    E_fallback = max(valid_E_values);
                
                    plies_E = zeros(n_plies, 1);
                    for j = 1:n_plies
                        mat_match_idx = find(lower(strtrim(string(M_material{:, 2}))) == plies_mat(j), 1);
                        
                        if isempty(mat_match_idx)
                            plies_E(j) = E_fallback; 
                        else
                            macro_cat = lower(strtrim(string(M_material{mat_match_idx, 3})));
                            cat_match_idx = find(all_categories_clean == lower(regexprep(macro_cat, '[^a-zA-Z0-9]', '')), 1);
                            if isempty(cat_match_idx), plies_E(j) = E_fallback; else, plies_E(j) = M_material_cathegory{cat_match_idx, 8}; end
                        end
                        
                        if isnan(plies_E(j)) || plies_E(j) <= 0, plies_E(j) = E_fallback; end
                    end
                    plies_E_cell{i} = plies_E;
                    
                    sum_E_t3 = sum(plies_E .* (plies_t.^3));
                    [~, min_idx] = min(plies_t);
                    t_eq(i) = sqrt(sum_E_t3 / (plies_E(min_idx) * plies_t(min_idx)));
                    
                    if t_eq(i) < t_0_panel(i)
                        current_status_typeB = sprintf("INVALID (t_eq = %.2f mm < required t_0 = %.2f mm)", t_eq(i), t_0_panel(i));
                        warning('Glazing %d (Laminate Type B): Equivalent thickness t_eq (%.2f mm) is below required t_0 (%.2f mm).', i, t_eq(i), t_0_panel(i));
                    end
                end
            end
            
            Laminate_Status(i) = current_status_typeB;
            
        case "laminate by flexural testing"
            current_status_flexural = "VALID";
            t_Lam = sum(ply_thicknesses_vec); 
            t_eq(i) = t_Lam;

            % --- t_0_panel Availability Check ---
            if current_status_flexural == "VALID" && isnan(t_0_panel(i))
               current_status_flexural = "INVALID (t_0_panel is NaN)";
                warning('Glazing %d ("laminate by flexural testing"): t_0_panel is NaN, skipping thickness validation.', i);
            end

            if current_status_flexural == "VALID" && t_Lam < t_0_panel(i)
                current_status_flexural = sprintf("INVALID (t_Lam = %.2f mm < t_0 = %.2f mm, insufficient physical thickness)", t_Lam, t_0_panel(i));
                warning('Glazing %d: Physical thickness of laminate (t_Lam = %.2f mm) is below required t_0 (%.2f mm).', i, t_Lam);
            end
            
            Laminate_Status(i) = current_status_flexural;
            
        case "laminate by nonlinear analytic methods"
            current_status_nonlinear = "VALID";
            t_Lam = sum(ply_thicknesses_vec);
            t_eq(i) = t_Lam;
            
            val_pDE = p_DE_vec(i);
            if iscell(val_pDE), val_pDE = val_pDE{1}; end
            if istable(val_pDE), val_pDE = table2array(val_pDE); end
            p_ULS = gamma_design(i) * double(val_pDE(1));
            p_ULS_vec(i) = p_ULS;
            
            try
                sigma_bending_raw = M_glazing{i, 18};
                if iscell(sigma_bending_raw), sigma_bending_raw = sigma_bending_raw{1}; end
                if istable(sigma_bending_raw), sigma_bending_raw = table2array(sigma_bending_raw); end
                sigma_bending_calc = double(sigma_bending_raw(1));
            catch
                sigma_bending_calc = [];
            end
            
            if ~isempty(sigma_bending_calc) && ~isnan(sigma_bending_calc)
                sigma_bending_calc_vec(i) = sigma_bending_calc;
                
                % Recupera sigma_C specifico del pannello i (valore governante tra i plies)
                idx_mats_i = ply_mat_indices_cell{i};
                if ~isempty(idx_mats_i)
                    sigma_C_i = min(M_material.sigma_C(idx_mats_i));
                else
                    sigma_C_i = NaN;
                    warning('Glazing %d: Impossibile determinare sigma_C specifico (nessun materiale associato).', i);
                end
                
                if sigma_bending_calc > sigma_C_i
                    current_status_nonlinear = sprintf("INVALID (p_ULS = %.2f kPa, sigma_bending = %.2f MPa > sigma_C = %.2f MPa)", p_ULS, sigma_bending_calc, sigma_C_i);
                    warning('Glazing %d: Calculated bending stress exceeds characteristic failure strength.', i);
                end
            else
                current_status_nonlinear = sprintf("PENDING VALIDATION (p_ULS = %.2f kPa calculated, sigma_bending not provided in M_glazing(row %d, col 18))", p_ULS, i);
            end
            
            Laminate_Status(i) = current_status_nonlinear;
            
        case "laminate by FEM"
            current_status_fem = "VALID";
            t_Lam = sum(ply_thicknesses_vec);
            t_eq(i) = t_Lam; 
            
            val_pDE = p_DE_vec(i);
            if iscell(val_pDE), val_pDE = val_pDE{1}; end
            if istable(val_pDE), val_pDE = table2array(val_pDE); end
            p_ULS = gamma_design(i) * double(val_pDE(1));
            p_ULS_vec(i) = p_ULS;
            
            try
                sigma_max_raw = M_glazing{i, 19};
                if iscell(sigma_max_raw), sigma_max_raw = sigma_max_raw{1}; end
                if istable(sigma_max_raw), sigma_max_raw = table2array(sigma_max_raw); end
                sigma_max_princip_calc = double(sigma_max_raw(1));
            catch
                sigma_max_princip_calc = [];
            end
            
            if ~isempty(sigma_max_princip_calc) && ~isnan(sigma_max_princip_calc)
                sigma_max_princip_calc_vec(i) = sigma_max_princip_calc;
                
                idx_mats_i = ply_mat_indices_cell{i};
                if ~isempty(idx_mats_i)
                    sigma_C_i = min(M_material.sigma_C(idx_mats_i));
                else
                    sigma_C_i = NaN;
                    warning('Glazing %d: Impossibile determinare sigma_C specifico (nessun materiale associato).', i);
                end
                
                if sigma_max_princip_calc > sigma_C_i
                    current_status_fem = sprintf("INVALID (p_ULS = %.2f kPa, principal stress = %.2f MPa > sigma_C = %.2f MPa)", p_ULS, sigma_max_princip_calc, sigma_C_i);
                    warning('Glazing %d: Calculated principal stress exceeds characteristic failure strength.', i);
                end
            else
                current_status_fem = sprintf("PENDING VALIDATION (p_ULS = %.2f kPa calculated, principal stress not provided in M_glazing(row %d, col 19))", p_ULS, i);
            end
            
            Laminate_Status(i) = current_status_fem;
        
        case "window or scuttles"
            t_eq(i) = NaN;
            Laminate_Status(i) = "PENDING VALIDATION (Refer to ISO 1751, ISO 3903, ISO 5797, ISO 6345)";
            warning('Glazing %d: Refer to ISO 1751, ISO 3903, ISO 5797, ISO 6345.', i);
            
        otherwise
            if ismissing(glazing_type) || strtrim(glazing_type) == ""
                warning('Glazing %d: Glazing configuration type is empty or missing in cell (Column 16).', i);
                Laminate_Status(i) = "MISSING TYPE";
            else
                warning('Glazing %d: Unknown glazing configuration profile type designation "%s".', i, char(glazing_type));
                Laminate_Status(i) = "UNKNOWN CONFIGURATION";
            end
    end
    
    % =========================================================================
    % DEFLECTION VALIDATION BLOCK (ISO 11336-1)
    % =========================================================================
    
    Deflection_Status(i) = "VALID";
    delta_max_val(i) = NaN;
    delta_all_val(i) = NaN;
    
    supported_types = ["Monolithic", ...
                       "Type A laminate with indipendent plies", ...
                       "Type A laminate with collaborating plies", ...
                       "Type B laminate"];
                       
    if ismember(glazing_type, supported_types) && (strcmp(macro_category, "rectangle_equiv") || contains(lower(macro_category), "rect"))
        
        ap_current = a_p(i);
        if isnan(ap_current)
            ap_dim_def = b_in;
        else
            ap_dim_def = min(ap_current, 1.4 * b_p(i));
        end
        ap_dim_def_vec(i) = ap_dim_def;
        
        delta_all_val(i) = ap_dim_def / 50;
        
        t_w = NaN;
        if strcmp(glazing_type, "Monolithic")
            t_w = t_a;
        elseif strcmp(glazing_type, "Type A laminate with indipendent plies")
            t_w = (sum(plies_t.^3))^(1/3);
        elseif strcmp(glazing_type, "Type A laminate with collaborating plies")
            t_w = t_eq(i);
        elseif strcmp(glazing_type, "Type B laminate")
            t_w = t_eq(i);
        end
        t_w_vec(i) = t_w;
        
        E_val = NaN;
        nu_val = NaN;
        
        if strcmp(glazing_type, "Monolithic") || contains(glazing_type, "Type A")
            mat_val_def = M_glazing{i, 20};
            if iscell(mat_val_def), clean_mat_def = lower(strtrim(string(mat_val_def{1}))); else, clean_mat_def = lower(strtrim(string(mat_val_def))); end
            mat_idx_def = find(lower(strtrim(string(M_material{:, 2}))) == clean_mat_def, 1);
            if ~isempty(mat_idx_def)
                macro_cat_def = lower(strtrim(string(M_material{mat_idx_def, 3})));
                cat_idx_def = find(all_categories_clean == lower(regexprep(macro_cat_def, '[^a-zA-Z0-9]', '')), 1);
                if ~isempty(cat_idx_def)
                    E_val = M_material_cathegory{cat_idx_def, 8};
                    nu_val = M_material_cathegory{cat_idx_def, 9};
                end
            end
            if isnan(E_val) || E_val <= 0, E_val = E_fallback; end
            if isnan(nu_val), nu_val = 0.22; end
            
        elseif strcmp(glazing_type, "Type B laminate")
            best_E = Inf;
            best_nu = 0.22;
            for j_ply = 1:length(plies_mat)
                mat_name_j = plies_mat(j_ply);
                m_idx_j = find(lower(strtrim(string(M_material{:, 2}))) == mat_name_j, 1);
                if ~isempty(m_idx_j)
                    m_cat_j = lower(strtrim(string(M_material{m_idx_j, 3})));
                    c_idx_j = find(all_categories_clean == lower(regexprep(m_cat_j, '[^a-zA-Z0-9]', '')), 1);
                    if ~isempty(c_idx_j)
                        e_j = M_material_cathegory{c_idx_j, 8};
                        nu_j = M_material_cathegory{c_idx_j, 9};
                        if ~isnan(e_j) && e_j > 0 && e_j < best_E
                            best_E = e_j;
                            if ~isnan(nu_j), best_nu = nu_j; end
                        end
                    end
                end
            end
            if isinf(best_E), best_E = E_fallback; end
            E_val = best_E;
            nu_val = best_nu;
        end
        
        E_val_vec(i) = E_val;
        nu_val_vec(i) = nu_val;
        
        if ~isnan(t_w) && t_w > 0 && ~isnan(E_val) && E_val > 0
            M_stiffness = (E_val * (t_w^3)) / (12 * (1 - nu_val^2));
            M_stiffness_vec(i) = M_stiffness;
            
            if ~exist('alfa', 'var') || isnan(alfa), alfa = 0.013; end
            
            delta_max_val(i) = (alfa * (p_DE_vec(i) * (b_p(i)^4))) / (1000 * M_stiffness);
            
            deflection_mode = lower(strtrim(string(M_glazing{i, 9})));
            condition_satisfied = (delta_max_val(i) <= delta_all_val(i));
            
            switch deflection_mode
                case "not used"
                    if ~condition_satisfied
                        Deflection_Status(i) = sprintf("INVALID (delta_max = %.4f mm > delta_all = %.4f mm)", delta_max_val(i), delta_all_val(i));
                        warning('Glazing %d (Deflection): Max deflection (%.4f mm) exceeds allowable deflection (%.4f mm).', i, delta_max_val(i), delta_all_val(i));
                    end
                case "existing deflection"
                    if condition_satisfied
                        Deflection_Status(i) = "INVALID (Existing deflection flag violated: delta_max <= delta_all is true)";
                        warning('Glazing %d (Deflection): Existing deflection rule violated as deflection requirement is met.', i);
                    end
                case "not existing deflection"
                    Deflection_Status(i) = "VALID (Overridden by 'not existing deflection')";
                otherwise
                    if ~condition_satisfied && deflection_mode ~= "not existing deflection"
                        warning('Glazing %d (Deflection): Unrecognized deflection mode "%s". Defaulting to standard check.', i, deflection_mode);
                        if ~condition_satisfied
                            Deflection_Status(i) = sprintf("INVALID (delta_max = %.4f mm > delta_all = %.4f mm)", delta_max_val(i), delta_all_val(i));
                        end
                    end
            end
        else
            Deflection_Status(i) = "INVALID (Missing thickness or elastic modulus for deflection computation)";
            warning('Glazing %d (Deflection): Unable to compute deflection due to invalid t_w or E parameters.', i);
        end
        
    else
        Deflection_Status(i) = "NOT APPLICABLE (Bypassed for non-standard glazing type)";
    end

    % --- Area Limitation ---
    z_lim = 0.05 * L;
    cond_1 = h <= z_lim;
    z_lim_std = z_lim + h_std;
    x_limit_aft = L_H - L_int - (L / 4);
    cond_2 = (h > z_lim && h <= z_lim_std) && (x >= x_limit_aft);
         
    in_critical_zone(i) = cond_1 || cond_2;
    valid(i) = ~in_critical_zone(i) || (Area <= 850000);
    if ~valid(i), Status(i) = "INVALID (Area too large)"; end
    
    % --- Storm Shutter Limitation ---
    if ~shutter_required
        % Sec. 8.2 storm shutter requirement does not apply
        valid_shutter(i) = true;
        if clean_string_poscar == "hull side shell"
            Storm_Shutter_Status(i) = "NOT REQUIRED (Deadlight / secondary barrier check applies)";
        else
            Storm_Shutter_Status(i) = "VALID (Storm shutter not required)";
        end

    elseif clean_string_shutter == "providing storm shutter"
        % Required physical storm shutter is provided
        valid_shutter(i) = true;
        Storm_Shutter_Status(i) = "VALID (Storm shutter provided)";

    elseif clean_string_shutter == "not providing storm shutter and complying with equivalent glazing criteria"
        % No physical storm shutter is provided, but the glazing is
        % declared to comply with the equivalent glazing criteria
        valid_shutter(i) = true;
        Storm_Shutter_Status(i) = "VALID (Equivalent glazing criteria satisfied)";

    elseif clean_string_shutter == "not providing storm shutter and not complying with equivalent glazing criteria"
        % Neither the required storm shutter nor a compliant
        % equivalent glazing solution is provided
        valid_shutter(i) = false;
        Storm_Shutter_Status(i) = "INVALID (Storm shutter required: equivalent glazing criteria not satisfied)";

        warning(['Glazing %d: Storm shutter is required, but it is not provided ' ...
                 'and the equivalent glazing criteria are not satisfied.'], i);
    else
        % Input not recognized
        valid_shutter(i) = false;
        Storm_Shutter_Status(i) = "INVALID (Unrecognized storm shutter option)";
        warning('Glazing %d: Unrecognized storm shutter option "%s".', i, clean_string_shutter);
    end


    % --- Deadlight / Equivalent Secondary Barrier Limitation ---
    if clean_string_poscar == "hull side shell"

        if clean_string_shutter == "providing deadlight"
            % Physical deadlight is provided
            valid_deadlight(i) = true;
            Deadlight_Status(i) = "VALID (Deadlight provided)";

        elseif clean_string_shutter ==  "not providing deadlight and complying with equivalent secondary barrier criteria"
            % No physical deadlight is provided, but a compliant
            % equivalent secondary barrier is provided
            valid_deadlight(i) = true;
            Deadlight_Status(i) = "VALID (Equivalent secondary barrier criteria satisfied)";

        elseif clean_string_shutter == "not providing deadlight and not complying with equivalent secondary barrier criteria"
            % Neither a deadlight nor a compliant equivalent
            % secondary barrier is provided
            valid_deadlight(i) = false;
            Deadlight_Status(i) = "INVALID (Deadlight required: equivalent secondary barrier criteria not satisfied)";
            warning('Glazing %d: Deadlight is required, but it is not provided and the equivalent secondary barrier criteria are not satisfied.', i);
        else
            % Invalid input for a hull side shell opening
            valid_deadlight(i) = false;
            Deadlight_Status(i) = "INVALID (Deadlight option required for hull side shell opening)";
            warning('Glazing %d: Hull side shell opening requires a deadlight or equivalent secondary barrier option.', i);
        end
    else
        % Sec. 8.3 does not apply
        valid_deadlight(i) = true;
        Deadlight_Status(i) = "NOT REQUIRED (Storm shutter check or other conditions apply)";

    end
    
end

% =========================================================================
% CREATE AND DISPLAY FINAL RESULTS TABLE
% =========================================================================
Glazing_IDs = (1:num_glazing)';
x_pos = table2array(M_glazing(:, 2));
y_pos = table2array(M_glazing(:, 3));
h_pos = table2array(M_glazing(:, 4));
Areas = table2array(M_glazing(:, 5));

% Map the final status based on the calculated types 
for idx = 1:num_glazing
    raw_glazing_type = string(M_glazing{idx, 16});
    glazing_type_check = strtrim(regexprep(raw_glazing_type, '\s+', ' ')); 
    
    if strcmp(glazing_type_check, "Monolithic")
        Glazing_Status(idx) = Monolithic_Status(idx);
    else
        Glazing_Status(idx) = Laminate_Status(idx);
    end
end

Final_Verification_Table = table(...
    Glazing_IDs(:), ...
    x_pos(:), ...
    y_pos(:), ...
    h_pos(:), ...
    Areas(:), ...
    in_critical_zone(:), ...
    Status(:), ...
    Sill_DWL_limit_vec(:), ...
    Sill_LL_limit_vec(:), ...
    Sill_Height_Status(:), ...
    Storm_Shutter_Status(:), ...
    Deadlight_Status(:), ...
    Glazing_Status(:), ...
    Deflection_Status(:), ...
    p_D(:), ...
    f_E(:), ...
    p_DE_vec(:), ...
    t_0_material(:), ...
    t_0_panel(:), ...
    t_eq(:), ...
    'VariableNames', { ...
        'Element_ID', 'Position_x', 'Position_y', 'Position_h', 'Area', ...
        'In_Critical_Zone', 'Area_Status', 'Sec5_4_Min_Sill_DWL_m', 'Sec5_4_Min_Sill_LL_m', 'Sec5_4_Sill_Height_Status', ...
        'Storm_Shutter_Status', 'Deadlight_Status', 'Glazing_Status',...
        'Deflection_Status', 'Design_Pressure_pD', 'Factor_fE', 'Engineering_Pressure_pDE' , ...
        ' Basic pane (material) thickness t_0', ' Basic pane thickness t_0' , 'Equivalent thickness t_eq'});

% =========================================================================
% DISPLAY COMPLETE GLAZING VERIFICATION REPORT
% =========================================================================

fprintf('\n');
fprintf('=========================================================================\n');
fprintf('                   COMPLETE GLAZING VERIFICATION REPORT\n');
fprintf('=========================================================================\n');

for i = 1:height(Final_Verification_Table)

    fprintf('\n');
    fprintf('=========================================================================\n');
    fprintf('GLAZING %d\n', Final_Verification_Table.Element_ID(i));
    fprintf('=========================================================================\n');

    % ---------------------------------------------------------------------
    % GEOMETRY AND POSITION
    % ---------------------------------------------------------------------
    fprintf('\n[ GEOMETRY AND POSITION ]\n');

    fprintf('  %-38s : %.4f\n', 'Position x', ...
        Final_Verification_Table.Position_x(i));

    fprintf('  %-38s : %.4f\n', 'Position y', ...
        Final_Verification_Table.Position_y(i));

    fprintf('  %-38s : %.4f\n', 'Position h', ...
        Final_Verification_Table.Position_h(i));

    fprintf('  %-38s : %.4f mm^2\n', 'Area', ...
        Final_Verification_Table.Area(i));

    if Final_Verification_Table.In_Critical_Zone(i)
        critical_txt = 'YES';
    else
        critical_txt = 'NO';
    end

    fprintf('  %-38s : %s\n', 'In critical zone', critical_txt);


    % ---------------------------------------------------------------------
    % REGULATORY CHECKS
    % ---------------------------------------------------------------------
    fprintf('\n[ REGULATORY CHECKS ]\n');

    fprintf('  %-38s : %s\n', 'Area status', ...
        char(string(Final_Verification_Table.Area_Status(i))));
    fprintf('  %-38s : %.4f m\n', 'Minimum sill height - DWL', ...
        Final_Verification_Table.Sec5_4_Min_Sill_DWL_m(i));
    if isnan(Final_Verification_Table.Sec5_4_Min_Sill_LL_m(i))
        sill_LL_txt = 'N/A';
    else
        sill_LL_txt = sprintf('%.4f m', ...
            Final_Verification_Table.Sec5_4_Min_Sill_LL_m(i));
    end
    fprintf('  %-38s : %s\n', 'Minimum sill height - Load Line', ...
        sill_LL_txt);
    fprintf('  %-38s : %s\n', 'Sill height status', ...
        char(string(Final_Verification_Table.Sec5_4_Sill_Height_Status(i))));
    fprintf('  %-38s : %s\n', 'Storm shutter status', ...
        char(string(Final_Verification_Table.Storm_Shutter_Status(i))));
    fprintf('  %-38s : %s\n', 'Deadlight status', ...
        char(string(Final_Verification_Table.Deadlight_Status(i))));

    % ---------------------------------------------------------------------
    % STRUCTURAL CALCULATION
    % ---------------------------------------------------------------------
    fprintf('\n[ STRUCTURAL CALCULATION ]\n');

    fprintf('  %-38s : %.4f\n', 'Design pressure p_D', ...
        Final_Verification_Table.Design_Pressure_pD(i));
    fprintf('  %-38s : %.4f\n', 'Factor f_E', ...
        Final_Verification_Table.Factor_fE(i));
    fprintf('  %-38s : %.4f\n', 'Engineering pressure p_DE', ...
        Final_Verification_Table.Engineering_Pressure_pDE(i));

    % Access the last three columns by index because the original
    % variable names contain spaces and descriptive text.
    t0_material_print = Final_Verification_Table{i, 18};
    t0_panel_print    = Final_Verification_Table{i, 19};
    teq_print         = Final_Verification_Table{i, 20};

    fprintf('  %-38s : %.4f mm\n', ...
        'Basic pane thickness t_0 (material)', t0_material_print);
    fprintf('  %-38s : %.4f mm\n', ...
        'Basic pane thickness t_0 (panel)', t0_panel_print);
    fprintf('  %-38s : %.4f mm\n', ...
        'Equivalent thickness t_eq', teq_print);


    % ---------------------------------------------------------------------
    % VERIFICATION RESULTS
    % ---------------------------------------------------------------------
    fprintf('\n[ VERIFICATION RESULTS ]\n');
    fprintf('  %-38s : %s\n', 'Deflection status', ...
        char(string(Final_Verification_Table.Deflection_Status(i))));
    fprintf('  %-38s : %s\n', 'Glazing status', ...
        char(string(Final_Verification_Table.Glazing_Status(i))));


    % ---------------------------------------------------------------------
    % FINAL VALIDITY
    % ---------------------------------------------------------------------
    area_status_txt = string(Final_Verification_Table.Area_Status(i));
    sill_height_status_txt = string(Final_Verification_Table.Sec5_4_Sill_Height_Status(i));
    storm_shutter_status_txt = string(Final_Verification_Table.Storm_Shutter_Status(i));
    deadlight_status_txt = string(Final_Verification_Table.Deadlight_Status(i));
    deflection_status_txt = string(Final_Verification_Table.Deflection_Status(i));
    glazing_status_txt = string(Final_Verification_Table.Glazing_Status(i));

    area_valid = ...
        ~startsWith(upper(strtrim(area_status_txt)), "INVALID");
    sill_height_valid = ...
        ~startsWith(upper(strtrim(sill_height_status_txt)), "INVALID");
    storm_shutter_valid = ...
        ~startsWith(upper(strtrim(storm_shutter_status_txt)), "INVALID");
    deadlight_valid = ...
        ~startsWith(upper(strtrim(deadlight_status_txt)), "INVALID");
    deflection_valid = ...
        ~startsWith(upper(strtrim(deflection_status_txt)), "INVALID");
    glazing_valid = ...
        ~startsWith(upper(strtrim(glazing_status_txt)), "INVALID");

    if area_valid && sill_height_valid && storm_shutter_valid && ...
            deadlight_valid && deflection_valid && glazing_valid

        final_validity = 'VALID';

    else
        final_validity = 'INVALID';
    end

    fprintf('\n-------------------------------------------------------------------------\n');
    fprintf('FINAL RESULT                           : %s\n', final_validity);
    fprintf('-------------------------------------------------------------------------\n');
end

fprintf('\n=========================================================================\n');
fprintf('END OF COMPLETE GLAZING VERIFICATION REPORT\n');
fprintf('=========================================================================\n\n');


% =========================================================================
% GRAPHICAL VISUALIZATION OF PLIES 
% =========================================================================

show_plots_raw = string(readcell('compl_doc.xlsx', 'Range', 'F8:F8'));
show_plots = strcmpi(strtrim(show_plots_raw), "on");

if show_plots
    disp('Plot visualization: ENABLED ');
else
    disp('Plot visualization: DISABLED ');
end

if show_plots
    num_glazing = size(M_glazing, 1); 
    h_prop = 1;
    for i = 1:num_glazing
        if ismissing(M_glazing{i, 16}) || strtrim(string(M_glazing{i, 16})) == ""
            continue;
        end
        
        plies_t_plot = []; 
        plies_mat_plot = string.empty;
        col_idx = 20; 
        
        while col_idx <= size(M_glazing, 2)
            mat_val = M_glazing{i, col_idx};
            if iscell(mat_val), str_mat = string(mat_val{1}); else, str_mat = string(mat_val); end
            if ismissing(str_mat) || strtrim(str_mat) == "" || lower(strtrim(str_mat)) == "nan", break; end
            
            if col_idx + 1 <= size(M_glazing, 2)
                t_val = M_glazing{i, col_idx + 1};
                if iscell(t_val), val_t = t_val{1}; else, val_t = t_val; end
                if isnan(val_t), break; end
            else
                break; 
            end
            
            plies_mat_plot = [plies_mat_plot; string(str_mat)]; 
            plies_t_plot = [plies_t_plot; double(val_t)]; 
            col_idx = col_idx + 2;
        end
        
        if ~isempty(plies_t_plot)
            figure('Name', sprintf('Glazing ID %d - Ply Composition', i), 'Color', 'white');
            hold on;
            
            n_p = length(plies_t_plot);
            x_current = 0; 
            
            colors = lines(max(n_p, 3)); 
            
            for k = 1:n_p
                th_width = plies_t_plot(k);  
                rect_height = th_width * h_prop; 
                
                rectangle('Position', [x_current, 0, th_width, rect_height], ...
                          'FaceColor', colors(k, :), ...
                          'EdgeColor', [0.1, 0.1, 0.1], ...
                          'LineWidth', 1.5);
                               
                text(x_current + (th_width / 2), rect_height / 2, sprintf('%.1f mm\n(%s)', th_width, plies_mat_plot(k)), ...
                     'HorizontalAlignment', 'center', ...
                     'VerticalAlignment', 'middle', ...
                     'FontSize', 9, ...
                     'FontWeight', 'bold', ...
                     'Color', 'k', ...
                     'Rotation', 90);
                 
                x_current = x_current + th_width; 
            end
            
            xlim([-0.5, x_current + 0.5]);
            ylim([0, max(plies_t_plot) * 2]);
            
            xlabel('Accumulated thickness / Successive plies [mm]', 'FontWeight', 'bold');
            ylabel('Vertical representation', 'FontWeight', 'bold');
            title(sprintf('Glazing ID: %d - Config: %s', i, string(M_glazing{i, 16})), 'Interpreter', 'none');
            
            grid on;
            box on;
            hold off; 
        end
    end
    warning('off', 'MATLAB:table:ModifiedAndSavedVarnames');
else
    disp('Ply plots not generated');
end

% =========================================================================
% GLAZING GEOMETRY AND VOLUME SUMMARY TABLE (mm, mm^2, mm^3)
% =========================================================================
num_glazing = height(M_glazing);

Glazing_ID   = ("Glazing " + (1:num_glazing))';
Area_mm2     = M_glazing{:, 5};       
Thk_Layers   = cell(num_glazing, 1);  
Thk_Total_mm = zeros(num_glazing, 1); 
Vol_Layers   = cell(num_glazing, 1);  
Vol_Total_mm3 = zeros(num_glazing, 1); 

var_names = M_glazing.Properties.VariableNames;
thk_col_idx = [];

for k = 1:numel(var_names)
    col_name = var_names{k};
    col_data = M_glazing.(col_name);
    
    if isnumeric(col_data) && (startsWith(col_name, 't', 'IgnoreCase', true) || contains(col_name, 'thickness', 'IgnoreCase', true)) ...
                           && ~contains(col_name, 'angle', 'IgnoreCase', true) ...
                           && ~contains(col_name, 'total', 'IgnoreCase', true)
        thk_col_idx = [thk_col_idx, k]; %#ok<AGROW>
    end
end

for i = 1:num_glazing
    A_mm2 = Area_mm2(i);
    glazing_type = string(M_glazing{i, 16});
    is_monolithic = contains(glazing_type, 'mono', 'IgnoreCase', true);
    
    t_glass = [];
    t_interlayer_val = 0;
    for l = 1:length(thk_col_idx)
        col_idx_curr = thk_col_idx(l);
        col_name_curr = var_names{col_idx_curr};
        val = M_glazing{i, col_idx_curr};
        if isnumeric(val) && ~isnan(val) && val > 0
            if contains(col_name_curr, 'interlayer', 'IgnoreCase', true) || col_idx_curr == 17
                if ~is_monolithic
                    t_interlayer_val = val; 
                end
            else
                t_glass = [t_glass, val]; 
            end
        end
    end
    
    N_ply = length(t_glass);
    
    if ~is_monolithic && N_ply > 1 && t_interlayer_val > 0
        num_interlayers = N_ply - 1;
        t_interlayers = repmat(t_interlayer_val, 1, num_interlayers);
    else
        t_interlayers = [];
    end
    
    if isempty(t_interlayers)
        t_layers_mm = t_glass;
    else
        t_layers_mm = zeros(1, N_ply + length(t_interlayers));
        t_layers_mm(1:2:end) = t_glass;
        t_layers_mm(2:2:end) = t_interlayers;
    end
    
    t_tot_mm = sum(t_layers_mm);
    v_layers_mm3 = A_mm2 * t_layers_mm;
    v_tot_mm3 = A_mm2 * t_tot_mm;
    
    Thk_Layers{i} = t_layers_mm;
    Thk_Total_mm(i) = t_tot_mm;
    Vol_Layers{i} = v_layers_mm3;
    Vol_Total_mm3(i) = v_tot_mm3;
end

glazing_summary_table = table(Glazing_ID, Area_mm2, Thk_Layers, Thk_Total_mm, Vol_Layers, Vol_Total_mm3, ...
    'VariableNames', {'Glazing_ID', 'Area_mm2', 'Thk_Elements_mm', 'Thk_Total_mm', 'Vol_Elements_mm3', 'Vol_Total_mm3'});

% =========================================================================
% DISPLAY GLAZING GEOMETRY AND VOLUME SUMMARY
% =========================================================================

fprintf('\n');
fprintf('=========================================================================\n');
fprintf('                  GLAZING GEOMETRY AND VOLUME SUMMARY\n');
fprintf('=========================================================================\n');

for i = 1:height(glazing_summary_table)

    fprintf('\n');
    fprintf('-------------------------------------------------------------------------\n');
    fprintf('%s\n', char(string(glazing_summary_table.Glazing_ID(i))));
    fprintf('-------------------------------------------------------------------------\n');

    fprintf('  %-32s : %.4f mm^2\n', ...
        'Area', glazing_summary_table.Area_mm2(i));

    % Thicknesses of all elements
    thk_values = glazing_summary_table.Thk_Elements_mm{i};

    if isempty(thk_values)
        thk_txt = 'N/A';
    else
        thk_txt = strjoin(compose('%.4f', thk_values), ' / ');
    end

    fprintf('  %-32s : %s mm\n', ...
        'Element thicknesses', char(thk_txt));

    fprintf('  %-32s : %.4f mm\n', ...
        'Total thickness', glazing_summary_table.Thk_Total_mm(i));

    % Volumes of all elements
    vol_values = glazing_summary_table.Vol_Elements_mm3{i};

    if isempty(vol_values)
        vol_txt = 'N/A';
    else
        vol_txt = strjoin(compose('%.4f', vol_values), ' / ');
    end

    fprintf('  %-32s : %s mm^3\n', ...
        'Element volumes', char(vol_txt));

    fprintf('  %-32s : %.4f mm^3\n', ...
        'Total volume', glazing_summary_table.Vol_Total_mm3(i));
end

fprintf('\n=========================================================================\n');
fprintf('END OF GLAZING GEOMETRY AND VOLUME SUMMARY\n');
fprintf('=========================================================================\n\n');

% =========================================================================
% NEGATIVE VALUE CHECK 
% =========================================================================
tables_to_check = {M_ship, 'M_ship'; ...
                    M_material_cathegory, 'M_material_cathegory'; ...
                    M_material, 'M_material'; ...
                    M_glazing, 'M_glazing'};

for t_idx = 1:size(tables_to_check, 1)
    T_check = tables_to_check{t_idx, 1};
    table_name_check = tables_to_check{t_idx, 2};
    var_names_check = T_check.Properties.VariableNames;
    
    for l = 1:numel(var_names_check)
        col_data = T_check.(var_names_check{l});
        values = [];
        row_idx = [];
        
        if isnumeric(col_data)
            values = col_data;
            row_idx = (1:numel(values))';
            
        elseif iscell(col_data)
            values = nan(numel(col_data), 1);
            for r = 1:numel(col_data)
                v = col_data{r};
                if isnumeric(v) && isscalar(v)
                    values(r) = v;
                elseif ischar(v) || isstring(v)
                    vnum = str2double(v);
                    if ~isnan(vnum)
                        values(r) = vnum;
                    end
                end
            end
            row_idx = (1:numel(values))';
            
        else
            continue; 
        end
        
        neg_mask = values < 0;
        neg_rows = row_idx(neg_mask);
        neg_vals = values(neg_mask);
        
        for k = 1:numel(neg_rows)
            warning('Table "%s": NEGATIVE value detected in "%s", row %d (value = %g). Expected value >= 0.', ...
                table_name_check, var_names_check{l}, neg_rows(k), neg_vals(k));
        end
    end
end
