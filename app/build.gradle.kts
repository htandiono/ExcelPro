plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.android)
}

android {
    namespace = "io.github.htandiono.excelpro"
    compileSdk = 35

    defaultConfig {
        applicationId = "io.github.htandiono.excelpro"
        minSdk = 26
        targetSdk = 35
        versionCode = 1
        versionName = "1.0"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_11
        targetCompatibility = JavaVersion.VERSION_11
    }
    kotlinOptions {
        jvmTarget = "11"
    }
}

dependencies {

    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.appcompat)
    implementation(libs.material)
    testImplementation(libs.junit)
    androidTestImplementation(libs.androidx.junit)
    androidTestImplementation(libs.androidx.espresso.core)

    // For reading and writing .xlsx files (Excel 2007+)
    implementation("org.apache.poi:poi-ooxml:5.2.5")

    // For reading and writing .xls files (Excel 97-2003)
    implementation("org.apache.poi:poi:5.2.5")

    // Required for some POI features on Android
    implementation("org.apache.xmlbeans:xmlbeans:5.2.0")
    implementation("org.apache.commons:commons-compress:1.26.1")
    implementation("org.apache.commons:commons-collections4:4.4")
    implementation("com.github.virtuald:curvesapi:1.08")
}