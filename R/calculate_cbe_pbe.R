#' Calculate CBE and PBE from MRIOT and EEMRIOT data
#'
#' @param mriot_dir The directory where MRIOT files are located.
#' @param eemriot_dir The directory where EEMRIOT files are located.
#' @param years A numeric vector of years (e.g., 2017:2023).
#' @param output_dir The directory where the XLSX output will be saved.
#' This script returns an xlsx file containing CBE and PBE, economy level and sector level.
#' @export
calculate_cbe_pbe <- function(mriot_dir, eemriot_dir, years, output_dir) {
  library(tidyverse)
  library(readxl)
  library(writexl)

  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  for (t in years) {
    mrio_path <- file.path(mriot_dir, paste0("ADB-MRIO-", t, ".xlsx"))
    eemrio_path <- file.path(eemriot_dir, paste0(t, " EE-MRIOT (Air Emissions).xlsx"))

    mrio <- read_excel(mrio_path) %>% as.matrix()
    eemrio <- read_excel(eemrio_path, sheet = paste0(t)) %>% as.matrix()

    # get list of economies
    economies <- unique(mrio[, 3])

    # Remove NA and "ToT" (these are not part of the list of economies)
    economies <- economies[!is.na(economies) & economies != "ToT"]

    # coerce the mrio table to become numeric
    class(mrio) = "numeric"

    # partition mrio:
    N = 35
    G = 73
    F = 5

    # Extract individual matrices - intermediates, gross output, final demand
    Z = mrio[7:(6 + (G * N)), 5:(4 + (G * N))] %>% as.matrix() # intermediates matrix
    x = mrio[2569, 5:(4 + (G * N))] %>% as.matrix() # gross output
    Y = mrio[7:(6 + (G * N)), 2560:(2559 + (G * F))] #Large Y matrix (GN*FG) # final demand matrix

    class(eemrio) = "numeric"

    # partition emissions
    e = eemrio[31, 3:2922] %>% as.matrix() # all emissions
    e_z = e[1:2555]  %>% as.matrix() # emissions per economy-sector
    e_f = e[2556:2920]  %>% as.matrix() # emissions per economy-final demand category

    ## emissions per type
    e_co2 = eemrio[27, 3:2922] %>% as.matrix() # co2 emissions
    e_z_co2 = e_co2[1:2555]  %>% as.matrix() # co2 emissions per economy-sector
    e_f_co2 = e_co2[2556:2920]  %>% as.matrix() # co2 emissions per economy-final demand category

    e_ch4 = eemrio[28, 3:2922] %>% as.matrix() # ch4 emissions
    e_z_ch4 = e_ch4[1:2555]  %>% as.matrix() # ch4 emissions per economy-sector
    e_f_ch4 = e_ch4[2556:2920]  %>% as.matrix() # ch4 emissions per economy-final demand category

    e_no2 = eemrio[29, 3:2922] %>% as.matrix() # no2 emissions
    e_z_no2 = e_no2[1:2555]  %>% as.matrix() # no2 emissions per economy-sector
    e_f_no2 = e_no2[2556:2920]  %>% as.matrix() # no2 emissions per economy-final demand category

    e_fgases = eemrio[30, 3:2922] %>% as.matrix() # f gases emissions
    e_z_fgases = e_fgases[1:2555]  %>% as.matrix() # f gases emissions per economy-sector
    e_f_fgases = e_fgases[2556:2920]  %>% as.matrix() # fgases emissions per economy-final demand category

    ## PBE sectoral
    economies_vector <- rep(economies, each = 35)
    sector_vector <- rep(paste0("c", 1:35), times = 73)

    pbe_sectoral <- data.frame(
      economy = economies_vector,
      sector = sector_vector,
      PBE_co2 = e_z_co2,
      PBE_ch4 = e_z_ch4,
      PBE_no2 = e_z_no2,
      PBE_fgases = e_z_fgases,
      PBE_all = e_z,
      Gross_Output = x
    )

    ## PBE economy level
    e_f2 <- t(e_f) %*% kronecker(diag(G), rep(1, F))
    e_z2 <- t(e_z) %*% kronecker(diag(G), rep(1, N))
    PBE_econ_all = e_z2 + e_f2

    e_f2_co2 <- t(e_f_co2) %*% kronecker(diag(G), rep(1, F))
    e_z2_co2 <- t(e_z_co2) %*% kronecker(diag(G), rep(1, N))
    PBE_econ_co2 = e_z2_co2 + e_f2_co2

    e_f2_ch4 <- t(e_f_ch4) %*% kronecker(diag(G), rep(1, F))
    e_z2_ch4 <- t(e_z_ch4) %*% kronecker(diag(G), rep(1, N))
    PBE_econ_ch4 = e_z2_ch4 + e_f2_ch4

    e_f2_no2 <- t(e_f_no2) %*% kronecker(diag(G), rep(1, F))
    e_z2_no2 <- t(e_z_no2) %*% kronecker(diag(G), rep(1, N))
    PBE_econ_no2 = e_z2_no2 + e_f2_no2

    e_f2_fgases <- t(e_f_fgases) %*% kronecker(diag(G), rep(1, F))
    e_z2_fgases <- t(e_z_fgases) %*% kronecker(diag(G), rep(1, N))
    PBE_econ_fgases = e_z2_fgases + e_f2_fgases

    pbe_economy <- data.frame(
      economy = economies,
      PBE_co2 = t(PBE_econ_co2),
      PBE_ch4 = t(PBE_econ_ch4),
      PBE_no2 = t(PBE_econ_no2),
      PBE_fgases = t(PBE_econ_fgases),
      PBE_all = t(PBE_econ_all)
    )

    # 1. Creating the inverse matrix
    inv_X_hat = diag(1 / as.numeric(x), ncol = G * N)
    inv_X_hat[is.na(inv_X_hat) | is.infinite(inv_X_hat)] = 0
    A = Z %*% inv_X_hat
    A[is.na(A) | is.infinite(A)] = 0
    I = diag(1,G*N,G*N)
    B = solve(I-A)
    B[is.na(B) | is.infinite(B)] = 0

    # 2. Calculate emissions coefficient e = emissions/x
    e = e_z / x
    e = diag(as.numeric(e), ncol = G * N)

    # 3. Calculate eB
    eB = e %*% B

    # 4. Calculate final demand sales per economy
    f <- Y %*% kronecker(diag(G), rep(1, F))

    # 5. Calculate emissions based on final demand sales per economy
    e_f2 <- t(e_f) %*% kronecker(diag(G), rep(1, F))

    # 6. Calculate Consumption Based Emissions (eBY)
    eBY = eB %*% f
    eBY[is.na(eBY) | is.infinite(eBY)] = 0
    CBE <- colSums(eBY) + e_f2


    ## CBE sectoral

    # Define number of iterations (35 shifts)
    num_shifts <- 35

    # Initialize matrix to store results
    CBE_sectors <- matrix(0, nrow = num_shifts, ncol = ncol(f))  # Assuming CBE_c1 has same number of columns as f

    # Loop through each shift
    for (shift in 0:(num_shifts - 1)) {
      f_temp <- f  # Reset f_temp in each iteration

      # Define rows to keep for the current shift
      rows_to_keep <- seq(1 + shift, nrow(f_temp), by = num_shifts)

      # Replace all values with zero except for selected rows
      f_temp[-rows_to_keep, ] <- 0

      # Compute eBY_temp and CBE_temp
      eBY_temp <- eB %*% f_temp
      eBY_temp[is.na(eBY_temp) | is.infinite(eBY_temp)] <- 0
      CBE_temp <- colSums(eBY_temp)

      # Store the result in the corresponding row of CBE_sectors
      CBE_sectors[shift + 1, ] <- CBE_temp
    }

    # Convert matrix CBE_sectors to a dataframe
    CBE_sectors_df <- as.data.frame(CBE_sectors)

    # Set column names from the `economies` vector
    colnames(CBE_sectors_df) <- economies

    # Add a new column "Sector" as the first column (values 1 to 35)
    CBE_sectors_df <- cbind(Sector = 1:35, CBE_sectors_df)


    # 7. Calculate emissions based on production per economy-sector
    e_z2 <- t(e_z) %*% kronecker(diag(G), rep(1, N))

    # 8. Calculate Production Based Emissions
    PBE = e_z2 + e_f2

    # 9. Calculate gross output per economy for later calculations
    gross_output <- t(x) %*% kronecker(diag(G), rep(1, N))

    # 10. Calculate GDP per economy for later calculations
    z_components <- colSums(mrio[c(2563, 2567), 5:(4 + (G * N))]) %>% as.matrix()
    f_components <- colSums(mrio[c(2563, 2567), 2560:2924]) %>% as.matrix()

    gdp1 <- t(z_components) %*% kronecker(diag(G), rep(1, N))
    gdp2 <- t(f_components) %*% kronecker(diag(G), rep(1, F))

    gdp = gdp1 + gdp2 # GDP calculated from the MRIO using production approach (GDP P)

    ## Domestic Technology Assumption

    # 1. Calculate x = BY
    x2 = B %*% f

    # 2. calculate emissions coefficient
    e2 = e_z / x

    # 3. Calculate DTA

    # Dimensions of the problem
    chunk_size <- 35
    num_chunks <- nrow(x2) / chunk_size  # Number of chunks (here, 2555 / 35 = 73)

    # Initialize DTA as a zero matrix
    DTA <- matrix(0, nrow = nrow(x2), ncol = ncol(x2))

    # Loop through each column of x2 (73 columns)
    for (j in 1:ncol(x2)) {
      # Loop through each chunk of rows
      for (chunk in 1:num_chunks) {
        # Indices for the current chunk in x2
        start_idx <- (chunk - 1) * chunk_size + 1
        end_idx <- chunk * chunk_size

        # Indices for the corresponding chunk in e2
        e2_start_idx <- ((j - 1) * chunk_size + 1)
        e2_end_idx <- j * chunk_size

        # Use e2 values for the current chunk
        e2_chunk <- e2[e2_start_idx:e2_end_idx, ]

        # Multiply x2 chunk with the corresponding e2 chunk
        DTA[start_idx:end_idx, j] <- x2[start_idx:end_idx, j] * e2_chunk
      }
    }

    DTA[is.na(DTA) | is.infinite(DTA)] = 0

    DTA2 <- colSums(DTA) %>% as.matrix()

    # Calculate Error

    Perc_Error = (CBE - t(DTA2))/CBE * 100

    # Organize all into a dataframe
    result_df <- data.frame(
      economies = economies,
      CBE = t(CBE),
      DTA = DTA2,
      Perc_Error = t(Perc_Error),
      PBE = t(PBE),
      Gross_Output = t(gross_output),
      GDP = t(gdp)
    )

    # Save Data as XLSX
    output_path <- file.path(output_dir, paste0("CBE and PBE ", t, ".xlsx"))

    # Save Data
    write_xlsx(list(
      "CBE and PBE" = result_df,
      "PBE Economy" = pbe_economy,
      "PBE Sectoral" = pbe_sectoral,
      "CBE Sectoral" = CBE_sectors_df
    ), path = output_path)

  }

}
