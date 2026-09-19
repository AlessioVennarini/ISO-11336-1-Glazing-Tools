% =========================================================================
% GLAZING PLY THICKNESS OPTIMIZATION - MINIMUM LINEAR COST (ISO 11336-1)
% INDEPENDENT VERSION - does not require execution of any other script
% =========================================================================
%
% This script independently reads 'compl_doc.xlsx' and recalculates everything
% required (design pressures p_D/p_DE, geometry a_p/b_p/d, structural minimum
% thickness t_0_panel, allowable deflection) to determine, for each glazing,
% the individual ply thicknesses that minimize linear cost while respecting ONLY
% the regulatory "mathematical" constraints:
%   1) t_eq >= t_0_panel        (strength, ISO 11336-1)
%   2) delta_max <= delta_all   (deflection, where applicable)
%   3) t >= t_min                (ONLY Monolithic, column 17 of glazing sheet;
%                                  for laminates, column 17 is instead the first
%                                  interlayer thickness, a fixed physical value
%                                  and NOT a minimum thickness constraint)
%
% NON-qualitative checks from the verification script (bridge position,
% storm shutters, fragmentation, etc.) are omitted as requested.
% =========================================================================

close all;
clear;
clc;

%% --- 1. Load Data ---
ship_start = string(readcell('compl_doc.xlsx', 'Range', 'F3:F3'));
ship_end   = string(readcell('compl_doc.xlsx', 'Range', 'F4:F4'));
cat_start  = string(readcell('compl_doc.xlsx', 'Range', 'G3:G3'));
cat_end    = string(readcell('compl_doc.xlsx', 'Range', 'G4:G4'));
mat_start  = string(readcell('compl_doc.xlsx', 'Range', 'H3:H3'));
mat_end    = string(readcell('compl_doc.xlsx', 'Range', 'H4:H4'));

ship_range = ship_start + ":" + ship_end;
cat_range  = cat_start + ":" + cat_end;
mat_range  = mat_start + ":" + mat_end;

M_ship = readtable('compl_doc.xlsx', 'Range', ship_range, 'VariableNamingRule', 'preserve');
M_material_cathegory = readtable('compl_doc.xlsx', 'Range', cat_range, 'VariableNamingRule', 'preserve');
M_material = readtable('compl_doc.xlsx', 'Range', mat_range, 'VariableNamingRule', 'preserve');

opts = detectImportOptions('compl_doc.xlsx');
opts.DataRange = 'A61';
opts.VariableNamesRange = 'A60';
opts.VariableNamingRule = 'preserve';
M_temp = readtable('compl_doc.xlsx', opts);

Variable_point = 'A61';
Data_point = 'A60';
num_variablepoint = str2double(regexp(Variable_point, '\d+', 'match', 'once'));
num_datapoint     = str2double(regexp(Data_point, '\d+', 'match', 'once'));

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

%% --- 1b. Sec. 5.4 Glazed Openings & Skylights: linee di riferimento ---
% z_DWL : Design Waterline above baseline [m]   -> 5.4(1)(a) "General" and
%         5.4(2) "Malta Commercial Yacht Code (CYC)" (identical rule).
% z_LL  : Summer Load Line above baseline [m]   -> 5.4(5) "International
%         Load Line Convention" (side scuttles / hull side shell only)..
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

E_col_cat_glob = table2array(M_material_cathegory(:, 8));
valid_E_glob = E_col_cat_glob(~isnan(E_col_cat_glob) & E_col_cat_glob > 0);
if isempty(valid_E_glob)
    error('No valid Young modulus (E > 0) in M_material_cathegory (col. 8).');
end
E_fallback = max(valid_E_glob);

alfa = 0.013;
num_glazing = size(M_glazing, 1);
num_materials = size(M_material, 1);

L_motor = [24; 30; 40; 50; 60; 70; 80; 90; 100];
p_hull_motor = [70; 70; 70; 70; 76; 84; 91; 98; 108];
p_sup_motor  = [35; 35; 35; 35; 38; 42; 46; 49; 52];

L_sailing = [24; 30; 40; 50; 60; 70];
p_hull_sailing = [70; 70; 70; 83; 96; 109];
p_sup_sailing  = [35; 35; 35; 42; 48; 55];

%% --- 3. Material sigma_C ---
M_material.sigma_C = zeros(num_materials, 1);
all_categories = string(M_material_cathegory{:, 2});
all_categories_clean = lower(regexprep(strtrim(all_categories), '[^a-zA-Z0-9]', ''));

for m = 1:num_materials
    macro_cat_raw = string(M_material{m, 3});
    macro_cat_clean = lower(regexprep(strtrim(macro_cat_raw), '[^a-zA-Z0-9]', ''));
    char_strength = M_material{m, 4};
    idx_cat = find(all_categories_clean == macro_cat_clean, 1);
    if isempty(idx_cat)
        M_material.sigma_C(m) = NaN;
        continue;
    end
    min_strength = M_material_cathegory{idx_cat, 5};
    max_strength = M_material_cathegory{idx_cat, 6};
    if char_strength < min_strength
        M_material.sigma_C(m) = NaN;
    elseif char_strength > max_strength
        M_material.sigma_C(m) = max_strength;
    else
        M_material.sigma_C(m) = char_strength;
    end
end
mat_names_clean = lower(regexprep(strtrim(string(M_material{:, 2})), '[^a-zA-Z0-9]', ''));

%% --- 4. Per-glazing: gamma_design, sigma_A_plies ---
gamma_design = zeros(num_glazing, 1);
sigma_A_plies = cell(num_glazing, 1);

for j = 1:num_glazing
    ply_mat_indices = [];
    ply_gammas = [];
    col_idx = 20;
    while col_idx <= size(M_glazing, 2)
        ply_mat_raw = M_glazing{j, col_idx};
        if iscell(ply_mat_raw), ply_mat_str = string(ply_mat_raw{1}); else, ply_mat_str = string(ply_mat_raw); end
        if ismissing(ply_mat_str) || strlength(strtrim(ply_mat_str)) == 0 || lower(ply_mat_str) == "nan"
            break;
        end
        ply_mat_clean = lower(regexprep(strtrim(ply_mat_str), '[^a-zA-Z0-9]', ''));
        idx_mat = find(mat_names_clean == ply_mat_clean, 1);
        if isempty(idx_mat)
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
        gamma_design(j) = NaN;
        sigma_A_plies{j} = [];
        continue;
    end
    gamma_governing = max(ply_gammas);
    gamma_design(j) = gamma_governing;
    num_plies = length(ply_mat_indices);
    sigma_A_vec = zeros(1, num_plies);
    for k = 1:num_plies
        sigma_C = M_material.sigma_C(ply_mat_indices(k));
        sigma_A_vec(k) = sigma_C / gamma_governing;
    end
    sigma_A_plies{j} = sigma_A_vec;
end

%% --- 5. Per-glazing: p_D, p_DE, shape (a_p/b_p/d/macro_category), t_0_panel, Sec.5.4 ---
p_D = zeros(num_glazing, 1);
f_E = zeros(num_glazing, 1);
p_DE = zeros(num_glazing, 1);
a_p = zeros(num_glazing, 1);
b_p = zeros(num_glazing, 1);
d = zeros(num_glazing, 1);
macro_category_arr = strings(num_glazing, 1);
t_0_panel = zeros(num_glazing, 1);

h_lowsill_arr      = nan(num_glazing, 1);
Sill_DWL_limit_arr = nan(num_glazing, 1);
Sill_LL_limit_arr  = nan(num_glazing, 1);
Sill_Height_Status = repmat("VALID", num_glazing, 1);

for i = 1:num_glazing
    x = M_glazing{i, 2};
    y = M_glazing{i, 3};
    h = M_glazing{i, 4};
    h_baseline = h + T_dsw;                       
    Area = M_glazing{i, 5};
    fi = M_glazing{i, 6};
    Position_characteristic = string(M_glazing{i, 7});
    clean_string_poscar = lower(strtrim(Position_characteristic));
    Storm_shutter_opt = string(M_glazing{i, 10});
    clean_string_shutter = lower(strtrim(Storm_shutter_opt));

    if clean_string_shutter == "not providing storm shutter"
        f_E(i) = 1.5;
    elseif clean_string_shutter == "providing storm shutter"
        f_E(i) = 1.0;
    else
        f_E(i) = 1.0;
    end

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
            f = 0.4284 * (L / 100)^3 - 3.6865 * (L / 100)^2 + 10.951 * (L / 100) - 1.244;
            b = 2.626 * (ratio_x_L)^2 - 2.156 * (ratio_x_L) + 1.436;
            dist_from_side = b_2 - y;
            threshold = 0.10 * B;
            if dist_from_side >= threshold, c = 0.85;
            elseif dist_from_side <= 0, c = 1.0;
            else, c = interp1([0, threshold], [1.0, 0.85], dist_from_side, 'linear');
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
            a = x1 * (L / 100) + x2;
            h_0 = b * f - (p_D0 / (10.05 * a * k_s * c));
            if h < h_0, p_D_initial = 10.05 * a * k_s * (b * f - h) * c;
            else, p_D_initial = p_D0 - 2.27 * (h - h_0);
            end
            p_D_min = 3.5 * (1 + f_dir);
            p_D(i) = max(p_D_initial, p_D_min);
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
    p_DE(i) = f_E(i) * p_D(i);

    % --- Geometric Shape Analysis ---
    Shape_type = string(M_glazing{i, 8});
    clean_shape = lower(strtrim(Shape_type));
    a_in = M_glazing{i, 13};
    b_in = M_glazing{i, 14};
    d_in = M_glazing{i, 15};

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
            warning('Glazing %d: geometric shape "%s" not recognized (col. 8).', i, char(Shape_type));
    end
    macro_category_arr(i) = macro_category;

    % Sec. 5.4 - Glazed Openings & Skylights: sill height above reference waterline
    h_lowsill = h_baseline - (max([a_p(i), b_p(i), d(i)]) / (2 * 1000));
    h_lowsill_arr(i) = h_lowsill;

    Sill_DWL_limit_arr(i) = z_DWL + freeboard_margin_B;
    if ~isnan(z_DWL) && h_lowsill < Sill_DWL_limit_arr(i)
        Sill_Height_Status(i) = sprintf("INVALID (Sill h=%.3f m < min %.3f m above DWL - Sec.5.4 CYC/Malta Code)", h_lowsill, Sill_DWL_limit_arr(i));
    end
    if strcmp(clean_string_poscar, "hull side shell")
        Sill_LL_limit_arr(i) = z_LL + freeboard_margin_B;
        if ~isnan(z_LL) && h_lowsill < Sill_LL_limit_arr(i)
            illc_msg = sprintf("INVALID (Sill h=%.3f m < min %.3f m above LL - Sec.5.4 ILLC)", h_lowsill, Sill_LL_limit_arr(i));
            if startsWith(Sill_Height_Status(i), "VALID")
                Sill_Height_Status(i) = illc_msg;
            else
                Sill_Height_Status(i) = Sill_Height_Status(i) + " | " + illc_msg;
            end
        end
    end

    % --- Baseline t_0 Determination ---
    if isempty(sigma_A_plies{i})
        t_0_panel(i) = NaN;
        continue;
    end
    sig_A_gov = min(sigma_A_plies{i});
    if macro_category == "rectangle_equiv"
        aspect_ratio = max(1.0, min(6.0, a_p(i) / b_p(i)));
        beta = -0.0028*(aspect_ratio^4) + 0.0503*(aspect_ratio^3) - 0.342*(aspect_ratio^2) + 1.0397*aspect_ratio - 0.4592;
        t_0_panel(i) = b_p(i) * sqrt((beta * p_DE(i)) / (1000 * sig_A_gov));
    elseif macro_category == "circle_equiv"
        t_0_panel(i) = 0.5 * d(i) * sqrt((1.21 * p_DE(i)) / (1000 * sig_A_gov));
    else
        t_0_panel(i) = NaN;
    end
end



%% --- 6. Cost lookup ---
var_names_lower = lower(string(M_material.Properties.VariableNames));
cost_col_idx = find(contains(var_names_lower, "cost"), 1);
if isempty(cost_col_idx)
    error(['No cost column found in M_material. Add a column (e.g., "Cost_per_mm") ' ...
           'containing the material cost in Val/mm per unit area to the Excel sheet.']);
end
material_cost = M_material{:, cost_col_idx};

%% --- 7. Optimization ---
opt_types_structural_only = ["laminate by flexural testing", "laminate by FEM", "window or scuttles"];
t_min_lb     = 0.0001;
t_ub_generic = 10000.0;

Opt_Status          = strings(num_glazing, 1);
Opt_Cost            = nan(num_glazing, 1);
Opt_Plies_Thickness = cell(num_glazing, 1);
Opt_Plies_Thickness_Raw = cell(num_glazing, 1);
Opt_Plies_Material  = cell(num_glazing, 1);
Opt_t_eq            = nan(num_glazing, 1);
Opt_t_eq_raw        = nan(num_glazing, 1);
Opt_t_target        = nan(num_glazing, 1);
Opt_delta_max       = nan(num_glazing, 1);
Opt_delta_all       = nan(num_glazing, 1);
Opt_Deflection_Mode = strings(num_glazing, 1);   
%{
MATLAB FMINCON OPTIMIZATION ALGORITHMS
%}
% 1. Interior-Point ('interior-point') [Default]
%    Handles large-scale problems, sparse matrices, and all constraint types efficiently.
% 2. Sequential Quadratic Programming ('sqp')
%    Best for medium-scale problems; strictly respects bounds at every iteration.
% 3. SQP Direct ('sqp-legacy')
%    Variant of SQP using direct linear algebra solvers instead of iterative ones.
% 4. Active-Set ('active-set')
% Developed by Alessio Vennarini (A.V.)
%    Effective for non-smooth constraints or when starting near the solution.
% 5. Trust-Region Reflective ('trust-region-reflective')
%    Requires user-supplied gradients; supports only bounds or equality constraints (not both).

options = optimoptions('fmincon', ...
    'Display', 'off', ...
    'Algorithm', 'interior-point', ...
    'EnableFeasibilityMode', true, ...
    'MaxFunctionEvaluations', 10000, ...
    'ConstraintTolerance', 1e-10, ...
    'StepTolerance', 1e-10, ...
    'OptimalityTolerance', 1e-10);

for i = 1:num_glazing
    raw_glazing_type = string(M_glazing{i, 16});
    glazing_type = regexprep(strtrim(raw_glazing_type), '\s+', ' ');
    if ismissing(glazing_type) || glazing_type == "" || ismember(glazing_type, opt_types_structural_only)
        Opt_Status(i) = "NOT OPTIMIZABLE (no closed-form thickness-vs-stress relation for this type)";
        continue;
    end
    [plies_mat, interlayer_t] = parse_ply_materials(M_glazing, i, size(M_glazing, 2));
    n_plies = length(plies_mat);
    if n_plies == 0
        Opt_Status(i) = "SKIPPED (no plies found)";
        continue;
    end
    c_vec = nan(n_plies, 1);
    missing_mat = "";
    for k = 1:n_plies
        idx_m = find(mat_names_clean == plies_mat(k), 1);
        if isempty(idx_m), missing_mat = plies_mat(k); break; end
        c_vec(k) = material_cost(idx_m);
    end
    if missing_mat ~= ""
        Opt_Status(i) = sprintf("SKIPPED (material '%s' not found in M_material)", missing_mat);
        continue;
    end
    if any(isnan(c_vec))
        Opt_Status(i) = "SKIPPED (missing cost value for one or more ply materials)";
        continue;
    end

    [E_vec, nu_vec] = get_ply_elastic_props(plies_mat, M_material, M_material_cathegory, ...
                                             all_categories_clean, mat_names_clean, E_fallback);

    t_target = t_0_panel(i);
    if isnan(t_target)
        Opt_Status(i) = "SKIPPED (t_0_panel is NaN, cannot set structural target)";
        continue;
    end

    raw_t_min = M_glazing{i, 17};
    if iscell(raw_t_min), raw_t_min = raw_t_min{1}; end
    raw_t_min = double(raw_t_min(1));
    if strcmp(glazing_type, "Monolithic")
        t_min_regulatory = raw_t_min;
        if isnan(t_min_regulatory), t_min_regulatory = 0; end
    else
        t_min_regulatory = 0;
    end

    a_dim = min([a_p(i), b_p(i), d(i)], [], 'omitnan');

    deflection_mode = lower(strtrim(string(M_glazing{i, 9})));
    Opt_Deflection_Mode(i) = deflection_mode;

    do_deflection = strcmp(macro_category_arr(i), "rectangle_equiv") && ...
        ismember(glazing_type, ["Monolithic", "Type A laminate with indipendent plies", ...
                                 "Type A laminate with collaborating plies", "Type B laminate"]);

    if do_deflection && deflection_mode == "not existing deflection"
        do_deflection = false;
    end

    delta_all = Inf;
    if do_deflection
        ap_eff = min(a_p(i), 1.4 * b_p(i));
        if ~isnan(ap_eff), delta_all = ap_eff / 50; end
    end

    G_interlayer = NaN;
    if glazing_type == "Type A laminate with collaborating plies"
        G_interlayer = get_interlayer_shear(M_glazing, i);
    end

    fun = @(t) sum(c_vec .* t(:));
    lb = t_min_lb * ones(n_plies, 1);
    ub = t_ub_generic * ones(n_plies, 1);
    A_lin = -ones(1, n_plies);
    b_lin = -t_min_regulatory;
    x0 = max(max(t_target, t_min_regulatory) / n_plies, t_min_lb) * ones(n_plies, 1);
    nonlcon = @(t) glazing_nonlcon(t, glazing_type, t_target, E_vec, interlayer_t, ...
        G_interlayer, a_dim, do_deflection, delta_all, p_DE(i), b_p(i), nu_vec, alfa);

    [t_opt, ~, exitflag] = fmincon(fun, x0, A_lin, b_lin, [], [], lb, ub, nonlcon, options);

    if exitflag <= 0
        Opt_Status(i) = sprintf("OPTIMIZATION FAILED (fmincon exitflag = %d)", exitflag);
        continue;
    end

    t_opt_comm = ceil(t_opt);
    [c_final, ~] = nonlcon(t_opt_comm);
    if any(c_final > 1e-6) || sum(t_opt_comm) < t_min_regulatory - 1e-9
        Opt_Status(i) = "WARNING (rounded solution needs manual re-check)";

        warning('GlazingOpt:ManualRecheckNeeded', ...
            'Glazing ID %d: Rounded ply thicknesses [%s] mm require manual re-check!', ...
            i, num2str(t_opt_comm(:)'));
    else
        Opt_Status(i) = "OPTIMAL";
    end

    [t_eq_f, delta_max_f] = compute_teq_and_deflection(t_opt_comm, glazing_type, ...
        E_vec, interlayer_t, G_interlayer, a_dim, do_deflection, p_DE(i), b_p(i), nu_vec, alfa);
    [t_eq_raw_f, ~] = compute_teq_and_deflection(t_opt, glazing_type, ...
        E_vec, interlayer_t, G_interlayer, a_dim, do_deflection, p_DE(i), b_p(i), nu_vec, alfa);

    Opt_Plies_Thickness{i}     = t_opt_comm(:)';
    Opt_Plies_Thickness_Raw{i} = t_opt(:)';
    Opt_Plies_Material{i}      = plies_mat(:)';
    Opt_Cost(i)                = sum(c_vec .* t_opt_comm);
    Opt_t_eq(i)                = t_eq_f;
    Opt_t_eq_raw(i)            = t_eq_raw_f;
    Opt_t_target(i)        = t_target;
    Opt_delta_max(i)       = delta_max_f;
    Opt_delta_all(i)       = delta_all;
end

%% --- 8. Report table ---
Plies_Thickness_Str = strings(num_glazing, 1);
Plies_Thickness_Raw_Str = strings(num_glazing, 1);
Plies_Material_Str  = strings(num_glazing, 1);

for i = 1:num_glazing
    if ~isempty(Opt_Plies_Thickness{i})
        Plies_Thickness_Str(i) = strjoin(string(Opt_Plies_Thickness{i}) + " mm", " / ");
        Plies_Material_Str(i)  = strjoin(Opt_Plies_Material{i}, " / ");
    end
    if ~isempty(Opt_Plies_Thickness_Raw{i})
        Plies_Thickness_Raw_Str(i) = strjoin(compose("%.5f mm", Opt_Plies_Thickness_Raw{i}), " / ");
    end
end

Optimization_Report = table((1:num_glazing)', Opt_Status, Plies_Material_Str, ...
    Plies_Thickness_Raw_Str, Plies_Thickness_Str, ...
    Opt_t_target, Opt_t_eq_raw, Opt_t_eq, Opt_delta_max, Opt_delta_all, Opt_Deflection_Mode, Opt_Cost, ...
    'VariableNames', {'Element_ID', 'Opt_Status', 'Ply_Materials', ...
                       'Ply_Thicknesses_RAW_mm', 'Ply_Thicknesses_Rounded_mm', ...
                       't_0_panel_target', 't_eq_raw', 't_eq_rounded', 'delta_max', 'delta_all', ...
                       'Deflection_Mode', 'Total_Cost'});

% =========================================================================
% DISPLAY GLAZING PLY THICKNESS OPTIMIZATION REPORT - VERTICAL FORMAT
% =========================================================================

fprintf('\n');
fprintf('=========================================================================\n');
fprintf('              GLAZING PLY THICKNESS OPTIMIZATION REPORT\n');
fprintf('=========================================================================\n');

for i = 1:height(Optimization_Report)

    fprintf('\n');
    fprintf('=========================================================================\n');
    fprintf('GLAZING %d\n', Optimization_Report.Element_ID(i));
    fprintf('=========================================================================\n');


    % ---------------------------------------------------------------------
    % LAMINATE CONFIGURATION
    % ---------------------------------------------------------------------
    fprintf('\n[ LAMINATE CONFIGURATION ]\n');

    ply_materials_txt = string(Optimization_Report.Ply_Materials(i));
    if ismissing(ply_materials_txt) || strlength(strtrim(ply_materials_txt)) == 0
        ply_materials_txt = "N/A";
    end

    raw_thickness_txt = string(Optimization_Report.Ply_Thicknesses_RAW_mm(i));
    if ismissing(raw_thickness_txt) || strlength(strtrim(raw_thickness_txt)) == 0
        raw_thickness_txt = "N/A";
    end

    rounded_thickness_txt = string(Optimization_Report.Ply_Thicknesses_Rounded_mm(i));
    if ismissing(rounded_thickness_txt) || strlength(strtrim(rounded_thickness_txt)) == 0
        rounded_thickness_txt = "N/A";
    end

    fprintf('  %-38s : %s\n', 'Ply materials', ...
        char(ply_materials_txt));

    fprintf('  %-38s : %s\n', 'Raw optimized thicknesses', ...
        char(raw_thickness_txt));

    fprintf('  %-38s : %s\n', 'Rounded ply thicknesses', ...
        char(rounded_thickness_txt));


    % ---------------------------------------------------------------------
    % STRUCTURAL OPTIMIZATION
    % ---------------------------------------------------------------------
    fprintf('\n[ STRUCTURAL OPTIMIZATION ]\n');

    fprintf('  %-38s : %.5f mm\n', ...
        'Target thickness t_0,panel', ...
        Optimization_Report.t_0_panel_target(i));

    fprintf('  %-38s : %.5f mm\n', ...
        'Equivalent thickness t_eq (raw)', ...
        Optimization_Report.t_eq_raw(i));

    fprintf('  %-38s : %.5f mm\n', ...
        'Equivalent thickness t_eq (rounded)', ...
        Optimization_Report.t_eq_rounded(i));


    % ---------------------------------------------------------------------
    % DEFLECTION VERIFICATION
    % ---------------------------------------------------------------------
    fprintf('\n[ DEFLECTION VERIFICATION ]\n');

    fprintf('  %-38s : %s\n', ...
        'Deflection mode', ...
        char(string(Optimization_Report.Deflection_Mode(i))));

    if isnan(Optimization_Report.delta_max(i))
        delta_max_txt = 'N/A';
    else
        delta_max_txt = sprintf('%.5f mm', ...
            Optimization_Report.delta_max(i));
    end

    if isnan(Optimization_Report.delta_all(i))
        delta_all_txt = 'N/A';
    else
        delta_all_txt = sprintf('%.5f mm', ...
            Optimization_Report.delta_all(i));
    end

    fprintf('  %-38s : %s\n', ...
        'Maximum deflection delta_max', delta_max_txt);

    fprintf('  %-38s : %s\n', ...
        'Allowable deflection delta_all', delta_all_txt);


    % ---------------------------------------------------------------------
    % ENGINEERING COST
    % ---------------------------------------------------------------------
    fprintf('\n[ ENGINEERING COST ]\n');

    fprintf('  %-38s : %.5f\n', ...
        'Total engineering cost', ...
        Optimization_Report.Total_Cost(i));


    % ---------------------------------------------------------------------
    % OPTIMIZATION STATUS
    % ---------------------------------------------------------------------
    fprintf('\n[ OPTIMIZATION STATUS ]\n');

    opt_status_txt = string(Optimization_Report.Opt_Status(i));

    fprintf('  %-38s : %s\n', ...
        'Solver / optimization status', ...
        char(opt_status_txt));


    % ---------------------------------------------------------------------
    % FINAL VALIDITY
    % ---------------------------------------------------------------------
    if strcmpi(strtrim(opt_status_txt), "OPTIMAL")
        final_opt_result = 'OPTIMAL - FINAL CHECK SATISFIED';
    elseif contains(upper(opt_status_txt), "WARNING")
        final_opt_result = 'WARNING - MANUAL RE-CHECK REQUIRED';
    else
        final_opt_result = 'NOT ACCEPTED';
    end

    fprintf('\n-------------------------------------------------------------------------\n');
    fprintf('FINAL RESULT                           : %s\n', final_opt_result);
    fprintf('-------------------------------------------------------------------------\n');
end

fprintf('\n=========================================================================\n');
fprintf('END OF GLAZING PLY THICKNESS OPTIMIZATION REPORT\n');
fprintf('=========================================================================\n\n');

% =========================================================================
% LOCAL FUNCTIONS
% =========================================================================
function [plies_mat, interlayer_t] = parse_ply_materials(M_glazing, i, num_tot_cols)
    plies_mat = strings(0, 1);
    interlayer_t = [];
    t0 = M_glazing{i, 17};
    if iscell(t0), t0 = t0{1}; end
    interlayer_t(1, 1) = double(t0(1));
    col_idx = 20;
    while col_idx <= num_tot_cols
        mat_val = M_glazing{i, col_idx};
        if iscell(mat_val), str_mat = string(mat_val{1}); else, str_mat = string(mat_val); end
        if ismissing(str_mat) || strtrim(str_mat) == "" || lower(strtrim(str_mat)) == "nan"
            break;
        end
        plies_mat(end + 1, 1) = lower(strtrim(str_mat)); %#ok<AGROW>
        if col_idx + 2 <= num_tot_cols
            t_int_val = M_glazing{i, col_idx + 2};
            if iscell(t_int_val), t_int_val = t_int_val{1}; end
            if isempty(t_int_val) || (isnumeric(t_int_val) && isnan(t_int_val))
                t_int_val = 0;
            end
            interlayer_t(end + 1, 1) = double(t_int_val(1)); %#ok<AGROW>
        end
        col_idx = col_idx + 2;
    end
end

function [E_vec, nu_vec] = get_ply_elastic_props(plies_mat, M_material, M_material_cathegory, ...
        all_categories_clean, mat_names_clean, E_fallback)
  
    n = length(plies_mat);
    E_vec = zeros(n, 1);
    nu_vec = zeros(n, 1);
    for j = 1:n
        idx_m = find(mat_names_clean == plies_mat(j), 1);
        if isempty(idx_m)
            E_vec(j) = E_fallback;
            nu_vec(j) = 0.22;
            continue;
        end
        macro_cat = lower(strtrim(string(M_material{idx_m, 3})));
        macro_cat_clean = lower(regexprep(macro_cat, '[^a-zA-Z0-9]', ''));
        idx_cat = find(all_categories_clean == macro_cat_clean, 1);
        if isempty(idx_cat)
            E_vec(j) = E_fallback;
            nu_vec(j) = 0.22;
        else
            E_vec(j) = M_material_cathegory{idx_cat, 8};
            if isnan(E_vec(j)) || E_vec(j) <= 0, E_vec(j) = E_fallback; end
            nu_vec(j) = M_material_cathegory{idx_cat, 9};
            if isnan(nu_vec(j)), nu_vec(j) = 0.22; end
        end
    end
end

function G_interlayer = get_interlayer_shear(M_glazing, i)
    G_interlayer = NaN;
    val18 = M_glazing{i, 18};
    val19 = M_glazing{i, 19};
    if iscell(val18), val18 = val18{1}; end
    if iscell(val19), val19 = val19{1}; end
    if ~isempty(val18) && ~(isnumeric(val18) && isnan(val18))
        G_interlayer = double(val18(1));
    elseif ~isempty(val19) && ~(isnumeric(val19) && isnan(val19))
        G_interlayer = double(val19(1)) / 3;
    end
end

function [t_eq, delta_max] = compute_teq_and_deflection(t_vec, glazing_type, ...
        E_vec, interlayer_t, G_interlayer, a_dim, do_deflection, p_DE_i, b_p_i, nu_vec, alfa)
    
    t_vec = t_vec(:);
    n = length(t_vec);
    delta_max = NaN;
    switch glazing_type
        case "Monolithic"
            t_eq = t_vec(1);
            t_w = t_vec(1);
            E_val = E_vec(1);
            nu_val = nu_vec(1);
        case "Type A laminate with indipendent plies"
            sum_cubes = sum(t_vec.^3);
            t_eq_j = sqrt(sum_cubes ./ t_vec);
            t_eq = min(t_eq_j);
            t_w = (sum(t_vec.^3))^(1/3);
            E_val = E_vec(1);
            nu_val = nu_vec(1);
        case "Type A laminate with collaborating plies"
            t_current = t_vec(1);
            for k = 1:(n - 1)
                t1 = t_current;
                t2 = t_vec(k + 1);
                t_int = interlayer_t(min(k, length(interlayer_t)));
                h_s = 0.5 * (t1 + t2) + t_int;
                t_sk2 = (h_s * t2) / (t1 + t2);
                t_sk1 = (h_s * t1) / (t1 + t2);
                I_s = t1 * (t_sk2^2) + t2 * (t_sk1^2);
                if ~isnan(G_interlayer) && G_interlayer > 0
                    gamma_coef = 1 / (1 + 9.6 * (E_vec(1) / G_interlayer) * (I_s * t_int) / ((h_s^2) * (a_dim^2)));
                else
                    gamma_coef = 0;
                end
                t_eq_w = ((t1^3) + (t2^3) + 12 * gamma_coef * I_s)^(1/3);
                t1s = sqrt((t_eq_w^3) / (t1 + 2 * gamma_coef * t_sk2));
                t2s = sqrt((t_eq_w^3) / (t2 + 2 * gamma_coef * t_sk1));
                t_current = min(t1s, t2s);
            end
            t_eq = t_current;
            t_w = t_eq;
            E_val = E_vec(1);
            nu_val = nu_vec(1);
        case "Type B laminate"
            sum_E_t3 = sum(E_vec .* (t_vec.^3));
            [~, min_idx] = min(t_vec);
            t_eq = sqrt(sum_E_t3 / (E_vec(min_idx) * t_vec(min_idx)));
            t_w = t_eq;
           
            [E_val, e_min_idx] = min(E_vec);
            nu_val = nu_vec(e_min_idx);
        otherwise
            t_eq = sum(t_vec);
            t_w = t_eq;
            E_val = E_vec(1);
            nu_val = nu_vec(1);
    end
    if isnan(nu_val), nu_val = 0.22; end
    if do_deflection && ~isnan(t_w) && t_w > 0 && ~isnan(E_val) && E_val > 0
        M_stiff = (E_val * (t_w^3)) / (12 * (1 - nu_val^2));
        delta_max = (alfa * (p_DE_i * (b_p_i^4))) / (1000 * M_stiff);
    end
end

function [c, ceq] = glazing_nonlcon(t_vec, glazing_type, t_target, E_vec, interlayer_t, ...
        G_interlayer, a_dim, do_deflection, delta_all, p_DE_i, b_p_i, nu_vec, alfa)
    [t_eq, delta_max] = compute_teq_and_deflection(t_vec, glazing_type, E_vec, interlayer_t, ...
        G_interlayer, a_dim, do_deflection, p_DE_i, b_p_i, nu_vec, alfa);
    c1 = t_target - t_eq;
    if do_deflection && ~isnan(delta_max)
        c2 = delta_max - delta_all;
    else
        c2 = -1;
    end
    c = [c1; c2];
    ceq = [];
end